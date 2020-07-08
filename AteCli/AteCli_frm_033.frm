VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_AutDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   1320
   ClientTop       =   390
   ClientWidth     =   12825
   Icon            =   "AteCli_frm_033.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10485
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   18494
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
         TabIndex        =   37
         Top             =   3540
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin TabDlg.SSTab tab_Princi 
            Height          =   4995
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   12615
            _ExtentX        =   22251
            _ExtentY        =   8811
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "AteCli_frm_033.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_NomGlo(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbl_NomGlo(2)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lbl_NomGlo(9)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lbl_NomGlo(8)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lbl_NomGlo(6)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lbl_NomGlo(19)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lbl_NomGlo(1)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lbl_NomGlo(3)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lbl_NomGlo(7)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "lbl_NomGlo(4)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "lbl_NomGlo(5)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SSPanel30"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_Tit_RegCyg"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_Tit_DirEle"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_Tit_Direcc"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_Tit_Telefo"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_Tit_Celula"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_Tit_LugNac"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "pnl_Tit_Profes"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "pnl_Tit_NivEst"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "pnl_Tit_EstCiv"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "pnl_Tit_Paises"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "pnl_Tit_FecNac"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "tab_DatCli"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).ControlCount=   24
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "AteCli_frm_033.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_NomGlo(26)"
            Tab(1).Control(1)=   "lbl_NomGlo(27)"
            Tab(1).Control(2)=   "lbl_NomGlo(29)"
            Tab(1).Control(3)=   "lbl_NomGlo(30)"
            Tab(1).Control(4)=   "lbl_NomGlo(31)"
            Tab(1).Control(5)=   "lbl_NomGlo(32)"
            Tab(1).Control(6)=   "lbl_NomGlo(35)"
            Tab(1).Control(7)=   "lbl_NomGlo(28)"
            Tab(1).Control(8)=   "lbl_NomGlo(34)"
            Tab(1).Control(9)=   "pnl_Cyg_ApeNom"
            Tab(1).Control(10)=   "pnl_Cyg_DocIde"
            Tab(1).Control(11)=   "SSPanel18"
            Tab(1).Control(12)=   "pnl_Cyg_DirEle"
            Tab(1).Control(13)=   "pnl_Cyg_Celula"
            Tab(1).Control(14)=   "pnl_Cyg_LugNac"
            Tab(1).Control(15)=   "pnl_Cyg_Profes"
            Tab(1).Control(16)=   "pnl_Cyg_NivEst"
            Tab(1).Control(17)=   "pnl_Cyg_Paises"
            Tab(1).Control(18)=   "pnl_Cyg_FecNac"
            Tab(1).Control(19)=   "tab_DatCyg"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Datos Crediticios"
            TabPicture(2)   =   "AteCli_frm_033.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lbl_NomGlo(36)"
            Tab(2).Control(1)=   "lbl_NomGlo(37)"
            Tab(2).Control(2)=   "lbl_NomGlo(38)"
            Tab(2).Control(3)=   "lbl_NomGlo(39)"
            Tab(2).Control(4)=   "lbl_NomGlo(40)"
            Tab(2).Control(5)=   "lbl_NomGlo(41)"
            Tab(2).Control(6)=   "lbl_NomGlo(42)"
            Tab(2).Control(7)=   "lbl_NomGlo(43)"
            Tab(2).Control(8)=   "lbl_NomGlo(44)"
            Tab(2).Control(9)=   "lbl_NomGlo(45)"
            Tab(2).Control(10)=   "lbl_NomGlo(46)"
            Tab(2).Control(11)=   "lbl_NomGlo(47)"
            Tab(2).Control(12)=   "lbl_NomGlo(48)"
            Tab(2).Control(13)=   "lbl_NomGlo(49)"
            Tab(2).Control(14)=   "lbl_NomGlo(50)"
            Tab(2).Control(15)=   "lbl_NomGlo(51)"
            Tab(2).Control(16)=   "lbl_NomGlo(52)"
            Tab(2).Control(17)=   "lbl_NomGlo(53)"
            Tab(2).Control(18)=   "lbl_NomGlo(54)"
            Tab(2).Control(19)=   "lbl_NomGlo(55)"
            Tab(2).Control(20)=   "lbl_NomGlo(190)"
            Tab(2).Control(21)=   "pnl_Cre_TasInt"
            Tab(2).Control(22)=   "pnl_Cre_CuoRen"
            Tab(2).Control(23)=   "SSPanel39"
            Tab(2).Control(24)=   "pnl_Cre_MonApr_MPr"
            Tab(2).Control(25)=   "pnl_Cre_MonApr_Sol"
            Tab(2).Control(26)=   "pnl_Cre_MonApr_Dol"
            Tab(2).Control(27)=   "pnl_Cre_TCaMPr"
            Tab(2).Control(28)=   "pnl_Cre_TCaDol"
            Tab(2).Control(29)=   "pnl_Cre_ILDCyg"
            Tab(2).Control(30)=   "pnl_Cre_ILDTit"
            Tab(2).Control(31)=   "pnl_Cre_PerGra"
            Tab(2).Control(32)=   "pnl_Cre_CuoExt"
            Tab(2).Control(33)=   "pnl_Cre_PlaApr"
            Tab(2).Control(34)=   "pnl_Cre_CuoFin"
            Tab(2).Control(35)=   "pnl_Cre_CuoIni"
            Tab(2).Control(36)=   "pnl_Cre_CuoFij"
            Tab(2).Control(37)=   "pnl_Cre_MonSol_MPr"
            Tab(2).Control(38)=   "pnl_Cre_MonSol_Sol"
            Tab(2).Control(39)=   "pnl_Cre_MonSol_Dol"
            Tab(2).Control(40)=   "pnl_Cre_ApoPro"
            Tab(2).Control(41)=   "pnl_Cre_TipMon"
            Tab(2).Control(42)=   "pnl_Cre_ComVta"
            Tab(2).ControlCount=   43
            TabCaption(3)   =   "Datos de Tasación"
            TabPicture(3)   =   "AteCli_frm_033.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txt_Tas_Observ"
            Tab(3).Control(1)=   "pnl_Tas_FecEmi"
            Tab(3).Control(2)=   "pnl_Tas_EmpPer"
            Tab(3).Control(3)=   "pnl_Tas_NumInf"
            Tab(3).Control(4)=   "pnl_Tas_FecEva"
            Tab(3).Control(5)=   "pnl_Tas_ValCom"
            Tab(3).Control(6)=   "pnl_Tas_NomPer"
            Tab(3).Control(7)=   "pnl_Tas_ValRea"
            Tab(3).Control(8)=   "pnl_Tas_AreCon"
            Tab(3).Control(9)=   "pnl_Tas_AreTer"
            Tab(3).Control(10)=   "SSPanel41"
            Tab(3).Control(11)=   "SSPanel42"
            Tab(3).Control(12)=   "SSPanel43"
            Tab(3).Control(13)=   "SSPanel44"
            Tab(3).Control(14)=   "pnl_Tas_VCoEs1"
            Tab(3).Control(15)=   "pnl_Tas_VReEs1"
            Tab(3).Control(16)=   "pnl_Tas_ACoEs1"
            Tab(3).Control(17)=   "pnl_Tas_ATeEs1"
            Tab(3).Control(18)=   "pnl_Tas_VCoEs2"
            Tab(3).Control(19)=   "pnl_Tas_VReEs2"
            Tab(3).Control(20)=   "pnl_Tas_ACoEs2"
            Tab(3).Control(21)=   "pnl_Tas_ATeEs2"
            Tab(3).Control(22)=   "pnl_Tas_VCoDep"
            Tab(3).Control(23)=   "pnl_Tas_VReDep"
            Tab(3).Control(24)=   "pnl_Tas_ACoDep"
            Tab(3).Control(25)=   "pnl_Tas_ATeDep"
            Tab(3).Control(26)=   "pnl_Tas_TotVCo"
            Tab(3).Control(27)=   "pnl_Tas_TotVRe"
            Tab(3).Control(28)=   "pnl_Tas_TotACo"
            Tab(3).Control(29)=   "pnl_Tas_TotATe"
            Tab(3).Control(30)=   "SSPanel45"
            Tab(3).Control(31)=   "SSPanel46"
            Tab(3).Control(32)=   "lbl_NomGlo(57)"
            Tab(3).Control(33)=   "lbl_NomGlo(66)"
            Tab(3).Control(34)=   "lbl_NomGlo(62)"
            Tab(3).Control(35)=   "lbl_NomGlo(59)"
            Tab(3).Control(36)=   "lbl_NomGlo(58)"
            Tab(3).Control(37)=   "lbl_NomGlo(61)"
            Tab(3).Control(38)=   "lbl_NomGlo(56)"
            Tab(3).Control(39)=   "lbl_NomGlo(60)"
            Tab(3).Control(40)=   "lbl_NomGlo(63)"
            Tab(3).Control(41)=   "lbl_NomGlo(64)"
            Tab(3).Control(42)=   "Line1"
            Tab(3).Control(43)=   "lbl_NomGlo(65)"
            Tab(3).ControlCount=   44
            TabCaption(4)   =   "Datos de Seguro"
            TabPicture(4)   =   "AteCli_frm_033.frx":007C
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "txt_Seg_ObsPol"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).Control(1)=   "txt_Seg_ObsEva"
            Tab(4).Control(1).Enabled=   0   'False
            Tab(4).Control(2)=   "pnl_Seg_PolTit"
            Tab(4).Control(2).Enabled=   0   'False
            Tab(4).Control(3)=   "pnl_Seg_EmiTit"
            Tab(4).Control(3).Enabled=   0   'False
            Tab(4).Control(4)=   "pnl_Seg_SegPre"
            Tab(4).Control(4).Enabled=   0   'False
            Tab(4).Control(5)=   "pnl_Seg_InfPre"
            Tab(4).Control(5).Enabled=   0   'False
            Tab(4).Control(6)=   "pnl_Seg_EvaPre"
            Tab(4).Control(6).Enabled=   0   'False
            Tab(4).Control(7)=   "pnl_Seg_AplPre"
            Tab(4).Control(7).Enabled=   0   'False
            Tab(4).Control(8)=   "pnl_Seg_FoiPre"
            Tab(4).Control(8).Enabled=   0   'False
            Tab(4).Control(9)=   "pnl_Seg_PolCyg"
            Tab(4).Control(9).Enabled=   0   'False
            Tab(4).Control(10)=   "pnl_Seg_PolViv"
            Tab(4).Control(10).Enabled=   0   'False
            Tab(4).Control(11)=   "pnl_Seg_EmiViv"
            Tab(4).Control(11).Enabled=   0   'False
            Tab(4).Control(12)=   "pnl_Seg_SegViv"
            Tab(4).Control(12).Enabled=   0   'False
            Tab(4).Control(13)=   "pnl_Seg_InfViv"
            Tab(4).Control(13).Enabled=   0   'False
            Tab(4).Control(14)=   "pnl_Seg_EvaViv"
            Tab(4).Control(14).Enabled=   0   'False
            Tab(4).Control(15)=   "pnl_Seg_AplViv"
            Tab(4).Control(15).Enabled=   0   'False
            Tab(4).Control(16)=   "pnl_Seg_FoiViv"
            Tab(4).Control(16).Enabled=   0   'False
            Tab(4).Control(17)=   "SSPanel47"
            Tab(4).Control(17).Enabled=   0   'False
            Tab(4).Control(18)=   "SSPanel49"
            Tab(4).Control(18).Enabled=   0   'False
            Tab(4).Control(19)=   "lbl_NomGlo(83)"
            Tab(4).Control(19).Enabled=   0   'False
            Tab(4).Control(20)=   "lbl_NomGlo(82)"
            Tab(4).Control(20).Enabled=   0   'False
            Tab(4).Control(21)=   "lbl_NomGlo(70)"
            Tab(4).Control(21).Enabled=   0   'False
            Tab(4).Control(22)=   "lbl_NomGlo(72)"
            Tab(4).Control(22).Enabled=   0   'False
            Tab(4).Control(23)=   "lbl_NomGlo(67)"
            Tab(4).Control(23).Enabled=   0   'False
            Tab(4).Control(24)=   "lbl_NomGlo(73)"
            Tab(4).Control(24).Enabled=   0   'False
            Tab(4).Control(25)=   "lbl_NomGlo(68)"
            Tab(4).Control(25).Enabled=   0   'False
            Tab(4).Control(26)=   "lbl_NomGlo(74)"
            Tab(4).Control(26).Enabled=   0   'False
            Tab(4).Control(27)=   "lbl_NomGlo(69)"
            Tab(4).Control(27).Enabled=   0   'False
            Tab(4).Control(28)=   "lbl_NomGlo(71)"
            Tab(4).Control(28).Enabled=   0   'False
            Tab(4).Control(29)=   "lbl_NomGlo(79)"
            Tab(4).Control(29).Enabled=   0   'False
            Tab(4).Control(30)=   "lbl_NomGlo(78)"
            Tab(4).Control(30).Enabled=   0   'False
            Tab(4).Control(31)=   "lbl_NomGlo(75)"
            Tab(4).Control(31).Enabled=   0   'False
            Tab(4).Control(32)=   "lbl_NomGlo(80)"
            Tab(4).Control(32).Enabled=   0   'False
            Tab(4).Control(33)=   "lbl_NomGlo(76)"
            Tab(4).Control(33).Enabled=   0   'False
            Tab(4).Control(34)=   "lbl_NomGlo(81)"
            Tab(4).Control(34).Enabled=   0   'False
            Tab(4).Control(35)=   "lbl_NomGlo(77)"
            Tab(4).Control(35).Enabled=   0   'False
            Tab(4).ControlCount=   36
            TabCaption(5)   =   "Datos Legales"
            TabPicture(5)   =   "AteCli_frm_033.frx":0098
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "txt_Leg_InfLeg"
            Tab(5).Control(1)=   "txt_Leg_ObsBlq"
            Tab(5).Control(2)=   "pnl_Leg_RepLeg"
            Tab(5).Control(3)=   "pnl_Leg_Notari"
            Tab(5).Control(4)=   "pnl_Leg_DocReg"
            Tab(5).Control(5)=   "pnl_Leg_AprCom"
            Tab(5).Control(6)=   "pnl_Leg_FirCon"
            Tab(5).Control(7)=   "pnl_Leg_FecBlq"
            Tab(5).Control(8)=   "lbl_NomGlo(84)"
            Tab(5).Control(9)=   "lbl_NomGlo(88)"
            Tab(5).Control(10)=   "lbl_NomGlo(87)"
            Tab(5).Control(11)=   "lbl_NomGlo(89)"
            Tab(5).Control(12)=   "lbl_NomGlo(91)"
            Tab(5).Control(13)=   "lbl_NomGlo(85)"
            Tab(5).Control(14)=   "lbl_NomGlo(86)"
            Tab(5).Control(15)=   "lbl_NomGlo(90)"
            Tab(5).ControlCount=   16
            TabCaption(6)   =   "Datos de COFIDE"
            TabPicture(6)   =   "AteCli_frm_033.frx":00B4
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "pnl_Cof_NumCar"
            Tab(6).Control(1)=   "pnl_Cof_NumOpe"
            Tab(6).Control(2)=   "pnl_Cof_FecEmi"
            Tab(6).Control(3)=   "pnl_Cof_FecVal"
            Tab(6).Control(4)=   "pnl_Cof_Import"
            Tab(6).Control(5)=   "pnl_Cof_TipMon"
            Tab(6).Control(6)=   "pnl_Cof_NomBan"
            Tab(6).Control(7)=   "pnl_Cof_NumCta"
            Tab(6).Control(8)=   "pnl_Cof_TasInt"
            Tab(6).Control(9)=   "pnl_Cof_TasCom"
            Tab(6).Control(10)=   "lbl_NomGlo(192)"
            Tab(6).Control(11)=   "lbl_NomGlo(191)"
            Tab(6).Control(12)=   "lbl_NomGlo(96)"
            Tab(6).Control(13)=   "lbl_NomGlo(95)"
            Tab(6).Control(14)=   "lbl_NomGlo(92)"
            Tab(6).Control(15)=   "lbl_NomGlo(97)"
            Tab(6).Control(16)=   "lbl_NomGlo(93)"
            Tab(6).Control(17)=   "lbl_NomGlo(94)"
            Tab(6).Control(18)=   "lbl_NomGlo(99)"
            Tab(6).Control(19)=   "lbl_NomGlo(98)"
            Tab(6).ControlCount=   20
            TabCaption(7)   =   "Datos del Inmueble"
            TabPicture(7)   =   "AteCli_frm_033.frx":00D0
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "lbl_NomGlo(106)"
            Tab(7).Control(1)=   "lbl_NomGlo(105)"
            Tab(7).Control(2)=   "lbl_NomGlo(104)"
            Tab(7).Control(3)=   "lbl_NomGlo(103)"
            Tab(7).Control(4)=   "lbl_NomGlo(102)"
            Tab(7).Control(5)=   "lbl_NomGlo(101)"
            Tab(7).Control(6)=   "lbl_NomGlo(33)"
            Tab(7).Control(7)=   "SSPanel5"
            Tab(7).Control(8)=   "SSPanel4"
            Tab(7).Control(9)=   "pnl_Inm_JurRep"
            Tab(7).Control(10)=   "pnl_Inm_JurDir"
            Tab(7).Control(11)=   "pnl_Inm_JurEmp"
            Tab(7).Control(12)=   "pnl_Inm_NatCyg"
            Tab(7).Control(13)=   "pnl_Inm_NatTit"
            Tab(7).Control(14)=   "pnl_Inm_TipPro"
            Tab(7).Control(15)=   "pnl_Inm_Direcc"
            Tab(7).ControlCount=   16
            Begin VB.TextBox txt_Leg_InfLeg 
               Height          =   1935
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   219
               Text            =   "AteCli_frm_033.frx":00EC
               Top             =   420
               Width           =   10755
            End
            Begin VB.TextBox txt_Leg_ObsBlq 
               Height          =   555
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   218
               Text            =   "AteCli_frm_033.frx":00F0
               Top             =   4350
               Width           =   10755
            End
            Begin VB.TextBox txt_Seg_ObsPol 
               Height          =   615
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   216
               Text            =   "AteCli_frm_033.frx":00F4
               Top             =   4320
               Width           =   10755
            End
            Begin VB.TextBox txt_Seg_ObsEva 
               Height          =   615
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   182
               Text            =   "AteCli_frm_033.frx":00F8
               Top             =   3690
               Width           =   10755
            End
            Begin VB.TextBox txt_Tas_Observ 
               Height          =   1125
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   139
               Top             =   3810
               Width           =   10725
            End
            Begin TabDlg.SSTab tab_DatCli 
               Height          =   2025
               Left            =   60
               TabIndex        =   42
               Top             =   2880
               Width           =   12435
               _ExtentX        =   21934
               _ExtentY        =   3572
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Actividad Principal"
               TabPicture(0)   =   "AteCli_frm_033.frx":00FC
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lbl_NomGlo(10)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "lbl_NomGlo(11)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "grd_Tit_ActPri"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "pnl_Tit_OcuPri"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "Actividad Secundaria"
               TabPicture(1)   =   "AteCli_frm_033.frx":0118
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "pnl_Tit_OcuSec"
               Tab(1).Control(1)=   "grd_Tit_ActSec"
               Tab(1).Control(2)=   "lbl_NomGlo(20)"
               Tab(1).Control(3)=   "lbl_NomGlo(21)"
               Tab(1).ControlCount=   4
               Begin Threed.SSPanel pnl_Tit_OcuPri 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   43
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Tit_ActPri 
                  Height          =   1245
                  Left            =   1680
                  TabIndex        =   44
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel pnl_Tit_OcuSec 
                  Height          =   315
                  Left            =   -73290
                  TabIndex        =   47
                  Top             =   390
                  Width           =   4215
                  _Version        =   65536
                  _ExtentX        =   7435
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Tit_ActSec 
                  Height          =   1245
                  Left            =   -73320
                  TabIndex        =   48
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   20
                  Left            =   -74880
                  TabIndex        =   50
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   21
                  Left            =   -74880
                  TabIndex        =   49
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   11
                  Left            =   120
                  TabIndex        =   46
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   10
                  Left            =   120
                  TabIndex        =   45
                  Top             =   390
                  Width           =   1275
               End
            End
            Begin Threed.SSPanel pnl_Cre_ComVta 
               Height          =   315
               Left            =   -73200
               TabIndex        =   38
               Top             =   750
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_TipMon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   39
               Top             =   420
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "NUEVOS SOLES"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_FecNac 
               Height          =   315
               Left            =   1800
               TabIndex        =   51
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Paises 
               Height          =   315
               Left            =   1800
               TabIndex        =   52
               Top             =   750
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_EstCiv 
               Height          =   315
               Left            =   1800
               TabIndex        =   53
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_NivEst 
               Height          =   315
               Left            =   1800
               TabIndex        =   54
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Profes 
               Height          =   315
               Left            =   8070
               TabIndex        =   55
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_LugNac 
               Height          =   315
               Left            =   1800
               TabIndex        =   56
               Top             =   1080
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Celula 
               Height          =   315
               Left            =   8070
               TabIndex        =   57
               Top             =   420
               Width           =   1875
               _Version        =   65536
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Telefo 
               Height          =   315
               Left            =   8070
               TabIndex        =   58
               Top             =   750
               Width           =   1875
               _Version        =   65536
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_Direcc 
               Height          =   615
               Left            =   1800
               TabIndex        =   59
               Top             =   2070
               Width           =   10725
               _Version        =   65536
               _ExtentX        =   18918
               _ExtentY        =   1085
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Tit_DirEle 
               Height          =   315
               Left            =   8070
               TabIndex        =   60
               Top             =   1080
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tit_RegCyg 
               Height          =   315
               Left            =   8070
               TabIndex        =   71
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel SSPanel30 
               Height          =   90
               Left            =   30
               TabIndex        =   73
               Top             =   2730
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin TabDlg.SSTab tab_DatCyg 
               Height          =   2025
               Left            =   -74940
               TabIndex        =   74
               Top             =   2880
               Width           =   12435
               _ExtentX        =   21934
               _ExtentY        =   3572
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Actividad Principal"
               TabPicture(0)   =   "AteCli_frm_033.frx":0134
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "lbl_NomGlo(24)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "lbl_NomGlo(25)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "grd_Cyg_ActPri"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "pnl_Cyg_OcuPri"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).ControlCount=   4
               TabCaption(1)   =   "Actividad Secundaria"
               TabPicture(1)   =   "AteCli_frm_033.frx":0150
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "pnl_Cyg_OcuSec"
               Tab(1).Control(1)=   "grd_Cyg_ActSec"
               Tab(1).Control(2)=   "lbl_NomGlo(108)"
               Tab(1).Control(3)=   "lbl_NomGlo(107)"
               Tab(1).ControlCount=   4
               Begin Threed.SSPanel pnl_Cyg_OcuPri 
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   75
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Cyg_ActPri 
                  Height          =   1245
                  Left            =   1680
                  TabIndex        =   76
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel pnl_Ocupac 
                  Height          =   315
                  Index           =   3
                  Left            =   -73290
                  TabIndex        =   77
                  Top             =   390
                  Width           =   4215
                  _Version        =   65536
                  _ExtentX        =   7435
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   1245
                  Index           =   3
                  Left            =   -73320
                  TabIndex        =   78
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel pnl_Cyg_OcuSec 
                  Height          =   315
                  Left            =   -73290
                  TabIndex        =   266
                  Top             =   390
                  Width           =   4095
                  _Version        =   65536
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _StockProps     =   15
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin MSFlexGridLib.MSFlexGrid grd_Cyg_ActSec 
                  Height          =   1245
                  Left            =   -73320
                  TabIndex        =   267
                  Top             =   720
                  Width           =   10665
                  _ExtentX        =   18812
                  _ExtentY        =   2196
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   108
                  Left            =   -74880
                  TabIndex        =   269
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   107
                  Left            =   -74880
                  TabIndex        =   268
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   25
                  Left            =   120
                  TabIndex        =   82
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   24
                  Left            =   120
                  TabIndex        =   81
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Ocupación:"
                  Height          =   285
                  Index           =   23
                  Left            =   -74880
                  TabIndex        =   80
                  Top             =   390
                  Width           =   1275
               End
               Begin VB.Label lbl_NomGlo 
                  Caption         =   "Datos Actividad:"
                  Height          =   285
                  Index           =   22
                  Left            =   -74880
                  TabIndex        =   79
                  Top             =   720
                  Width           =   1275
               End
            End
            Begin Threed.SSPanel pnl_Cyg_FecNac 
               Height          =   315
               Left            =   -73200
               TabIndex        =   83
               Top             =   1080
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_Paises 
               Height          =   315
               Left            =   -73200
               TabIndex        =   84
               Top             =   1410
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_NivEst 
               Height          =   315
               Left            =   -73200
               TabIndex        =   85
               Top             =   2070
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_Profes 
               Height          =   315
               Left            =   -66930
               TabIndex        =   86
               Top             =   2070
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_LugNac 
               Height          =   315
               Left            =   -73200
               TabIndex        =   87
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_Celula 
               Height          =   315
               Left            =   -66930
               TabIndex        =   88
               Top             =   1080
               Width           =   1875
               _Version        =   65536
               _ExtentX        =   3307
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_DirEle 
               Height          =   315
               Left            =   -66930
               TabIndex        =   89
               Top             =   1740
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   90
               Left            =   -74970
               TabIndex        =   97
               Top             =   2730
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel pnl_Cyg_DocIde 
               Height          =   315
               Left            =   -73200
               TabIndex        =   98
               Top             =   420
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cyg_ApeNom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   100
               Top             =   750
               Width           =   10725
               _Version        =   65536
               _ExtentX        =   18918
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cre_ApoPro 
               Height          =   315
               Left            =   -73200
               TabIndex        =   102
               Top             =   1080
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonSol_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   104
               Top             =   1410
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonSol_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   106
               Top             =   1410
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonSol_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   108
               Top             =   1410
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_CuoFij 
               Height          =   315
               Left            =   -73200
               TabIndex        =   110
               Top             =   2220
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_CuoIni 
               Height          =   315
               Left            =   -73200
               TabIndex        =   112
               Top             =   2550
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_CuoFin 
               Height          =   315
               Left            =   -73200
               TabIndex        =   114
               Top             =   2880
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_PlaApr 
               Height          =   315
               Left            =   -73200
               TabIndex        =   116
               Top             =   3210
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_CuoExt 
               Height          =   315
               Left            =   -73200
               TabIndex        =   118
               Top             =   3540
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cre_PerGra 
               Height          =   315
               Left            =   -73200
               TabIndex        =   120
               Top             =   3870
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_ILDTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   122
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_ILDCyg 
               Height          =   315
               Left            =   -68970
               TabIndex        =   124
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_TCaDol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   126
               Top             =   4530
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_TCaMPr 
               Height          =   315
               Left            =   -68970
               TabIndex        =   128
               Top             =   4530
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonApr_Dol 
               Height          =   315
               Left            =   -73200
               TabIndex        =   130
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonApr_Sol 
               Height          =   315
               Left            =   -68970
               TabIndex        =   132
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MonApr_MPr 
               Height          =   315
               Left            =   -64830
               TabIndex        =   134
               Top             =   1890
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel39 
               Height          =   90
               Left            =   -74970
               TabIndex        =   136
               Top             =   1770
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel pnl_Cre_CuoRen 
               Height          =   315
               Left            =   -64830
               TabIndex        =   137
               Top             =   4200
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_FecEmi 
               Height          =   315
               Left            =   -66120
               TabIndex        =   140
               Top             =   420
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tas_EmpPer 
               Height          =   315
               Left            =   -73200
               TabIndex        =   141
               Top             =   420
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tas_NumInf 
               Height          =   315
               Left            =   -73200
               TabIndex        =   142
               Top             =   750
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tas_FecEva 
               Height          =   315
               Left            =   -66120
               TabIndex        =   143
               Top             =   750
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tas_ValCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   144
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_NomPer 
               Height          =   315
               Left            =   -73200
               TabIndex        =   145
               Top             =   1080
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Tas_ValRea 
               Height          =   315
               Left            =   -71520
               TabIndex        =   146
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_AreCon 
               Height          =   315
               Left            =   -68160
               TabIndex        =   147
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_AreTer 
               Height          =   315
               Left            =   -69840
               TabIndex        =   148
               Top             =   1890
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -73200
               TabIndex        =   149
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Valor Comerc. US$"
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -71520
               TabIndex        =   150
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Valor Fabricac. US$"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -69840
               TabIndex        =   151
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Area Terreno m2"
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -68160
               TabIndex        =   152
               Top             =   1560
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Area Constr. m2"
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
            Begin Threed.SSPanel pnl_Tas_VCoEs1 
               Height          =   315
               Left            =   -73200
               TabIndex        =   153
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_VReEs1 
               Height          =   315
               Left            =   -71520
               TabIndex        =   154
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ACoEs1 
               Height          =   315
               Left            =   -68160
               TabIndex        =   155
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ATeEs1 
               Height          =   315
               Left            =   -69840
               TabIndex        =   156
               Top             =   2220
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_VCoEs2 
               Height          =   315
               Left            =   -73200
               TabIndex        =   157
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_VReEs2 
               Height          =   315
               Left            =   -71520
               TabIndex        =   158
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ACoEs2 
               Height          =   315
               Left            =   -68160
               TabIndex        =   159
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ATeEs2 
               Height          =   315
               Left            =   -69840
               TabIndex        =   160
               Top             =   2550
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_VCoDep 
               Height          =   315
               Left            =   -73200
               TabIndex        =   161
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_VReDep 
               Height          =   315
               Left            =   -71520
               TabIndex        =   162
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ACoDep 
               Height          =   315
               Left            =   -68160
               TabIndex        =   163
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_ATeDep 
               Height          =   315
               Left            =   -69840
               TabIndex        =   164
               Top             =   2880
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_TotVCo 
               Height          =   315
               Left            =   -73200
               TabIndex        =   165
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_TotVRe 
               Height          =   315
               Left            =   -71520
               TabIndex        =   166
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_TotACo 
               Height          =   315
               Left            =   -68160
               TabIndex        =   167
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Tas_TotATe 
               Height          =   315
               Left            =   -69840
               TabIndex        =   168
               Top             =   3330
               Width           =   1635
               _Version        =   65536
               _ExtentX        =   2884
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel45 
               Height          =   90
               Left            =   -74970
               TabIndex        =   180
               Top             =   1440
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel SSPanel46 
               Height          =   90
               Left            =   -74970
               TabIndex        =   181
               Top             =   3690
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel pnl_Seg_PolTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   183
               Top             =   1410
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_EmiTit 
               Height          =   315
               Left            =   -66150
               TabIndex        =   184
               Top             =   1410
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
               ForeColor       =   32768
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
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_Seg_SegPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   185
               Top             =   420
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_InfPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   186
               Top             =   750
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_EvaPre 
               Height          =   315
               Left            =   -66150
               TabIndex        =   187
               Top             =   750
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
               ForeColor       =   32768
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
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_Seg_AplPre 
               Height          =   315
               Left            =   -73200
               TabIndex        =   188
               Top             =   1080
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_FoiPre 
               Height          =   315
               Left            =   -66150
               TabIndex        =   189
               Top             =   1080
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Seg_PolCyg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   190
               Top             =   1740
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_PolViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   191
               Top             =   3210
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_EmiViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   192
               Top             =   3210
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
               ForeColor       =   32768
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
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_Seg_SegViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   193
               Top             =   2220
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_InfViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   194
               Top             =   2550
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_EvaViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   195
               Top             =   2550
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "01/10/2004"
               ForeColor       =   32768
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
               Font3D          =   2
            End
            Begin Threed.SSPanel pnl_Seg_AplViv 
               Height          =   315
               Left            =   -73200
               TabIndex        =   196
               Top             =   2880
               Width           =   2655
               _Version        =   65536
               _ExtentX        =   4683
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Seg_FoiViv 
               Height          =   315
               Left            =   -66150
               TabIndex        =   197
               Top             =   2880
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.9812345 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel47 
               Height          =   90
               Left            =   -74970
               TabIndex        =   214
               Top             =   2100
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   90
               Left            =   -74970
               TabIndex        =   215
               Top             =   3570
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel pnl_Leg_RepLeg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   220
               Top             =   3360
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Leg_Notari 
               Height          =   315
               Left            =   -73200
               TabIndex        =   221
               Top             =   3030
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Leg_DocReg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   222
               Top             =   4020
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Leg_AprCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   223
               Top             =   2370
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Leg_FirCon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   224
               Top             =   2700
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Leg_FecBlq 
               Height          =   315
               Left            =   -73200
               TabIndex        =   225
               Top             =   3690
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_NumCar 
               Height          =   315
               Left            =   -73200
               TabIndex        =   234
               Top             =   420
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_NumOpe 
               Height          =   315
               Left            =   -73200
               TabIndex        =   235
               Top             =   2070
               Width           =   2445
               _Version        =   65536
               _ExtentX        =   4313
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_FecEmi 
               Height          =   315
               Left            =   -73200
               TabIndex        =   236
               Top             =   750
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_FecVal 
               Height          =   315
               Left            =   -73200
               TabIndex        =   237
               Top             =   1080
               Width           =   1155
               _Version        =   65536
               _ExtentX        =   2037
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_Import 
               Height          =   315
               Left            =   -73200
               TabIndex        =   238
               Top             =   2730
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cof_TipMon 
               Height          =   315
               Left            =   -73200
               TabIndex        =   239
               Top             =   2400
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "NUEVOS SOLES"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_NomBan 
               Height          =   315
               Left            =   -73200
               TabIndex        =   246
               Top             =   1410
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Cof_NumCta 
               Height          =   315
               Left            =   -73200
               TabIndex        =   248
               Top             =   1740
               Width           =   3315
               _Version        =   65536
               _ExtentX        =   5847
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "31/12/2004"
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Inm_Direcc 
               Height          =   615
               Left            =   -73200
               TabIndex        =   250
               Top             =   420
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   1085
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Inm_TipPro 
               Height          =   315
               Left            =   -73200
               TabIndex        =   251
               Top             =   1050
               Width           =   3825
               _Version        =   65536
               _ExtentX        =   6747
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Inm_NatTit 
               Height          =   315
               Left            =   -73200
               TabIndex        =   252
               Top             =   1530
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Inm_NatCyg 
               Height          =   315
               Left            =   -73200
               TabIndex        =   253
               Top             =   1860
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Inm_JurEmp 
               Height          =   315
               Left            =   -73200
               TabIndex        =   254
               Top             =   2340
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_Inm_JurDir 
               Height          =   615
               Left            =   -73200
               TabIndex        =   255
               Top             =   2670
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   1085
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Inm_JurRep 
               Height          =   315
               Left            =   -73200
               TabIndex        =   256
               Top             =   3300
               Width           =   10755
               _Version        =   65536
               _ExtentX        =   18971
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   90
               Left            =   -74970
               TabIndex        =   264
               Top             =   1410
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   90
               Left            =   -74970
               TabIndex        =   265
               Top             =   2220
               Width           =   12525
               _Version        =   65536
               _ExtentX        =   22093
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
            Begin Threed.SSPanel pnl_Cre_TasInt 
               Height          =   315
               Left            =   -64830
               TabIndex        =   274
               Top             =   420
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cof_TasInt 
               Height          =   315
               Left            =   -73200
               TabIndex        =   276
               Top             =   3060
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cof_TasCom 
               Height          =   315
               Left            =   -73200
               TabIndex        =   278
               Top             =   3390
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa Comisión:"
               Height          =   315
               Index           =   192
               Left            =   -74880
               TabIndex        =   279
               Top             =   3420
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa Interés:"
               Height          =   315
               Index           =   191
               Left            =   -74880
               TabIndex        =   277
               Top             =   3090
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tasa de Interés:"
               Height          =   315
               Index           =   190
               Left            =   -66510
               TabIndex        =   275
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   33
               Left            =   -74880
               TabIndex        =   263
               Top             =   420
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Propietario:"
               Height          =   315
               Index           =   101
               Left            =   -74880
               TabIndex        =   262
               Top             =   1050
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Titular:"
               Height          =   315
               Index           =   102
               Left            =   -74880
               TabIndex        =   261
               Top             =   1530
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cónyuge:"
               Height          =   315
               Index           =   103
               Left            =   -74880
               TabIndex        =   260
               Top             =   1860
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa:"
               Height          =   315
               Index           =   104
               Left            =   -74880
               TabIndex        =   259
               Top             =   2340
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección Empresa:"
               Height          =   285
               Index           =   105
               Left            =   -74880
               TabIndex        =   258
               Top             =   2670
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Rep. Legal:"
               Height          =   315
               Index           =   106
               Left            =   -74880
               TabIndex        =   257
               Top             =   3300
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número de Cuenta:"
               Height          =   315
               Index           =   96
               Left            =   -74880
               TabIndex        =   249
               Top             =   1740
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Banco Transferencia:"
               Height          =   315
               Index           =   95
               Left            =   -74880
               TabIndex        =   247
               Top             =   1410
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Carta:"
               Height          =   285
               Index           =   92
               Left            =   -74880
               TabIndex        =   245
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Operación:"
               Height          =   285
               Index           =   97
               Left            =   -74880
               TabIndex        =   244
               Top             =   2070
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Carta:"
               Height          =   315
               Index           =   93
               Left            =   -74880
               TabIndex        =   243
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión Valor:"
               Height          =   315
               Index           =   94
               Left            =   -74880
               TabIndex        =   242
               Top             =   1080
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Importe:"
               Height          =   315
               Index           =   99
               Left            =   -74880
               TabIndex        =   241
               Top             =   2760
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda:"
               Height          =   315
               Index           =   98
               Left            =   -74880
               TabIndex        =   240
               Top             =   2400
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Informe Legal:"
               Height          =   315
               Index           =   84
               Left            =   -74880
               TabIndex        =   233
               Top             =   420
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Repres. Legal (es):"
               Height          =   315
               Index           =   88
               Left            =   -74880
               TabIndex        =   232
               Top             =   3360
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Notaria:"
               Height          =   315
               Index           =   87
               Left            =   -74880
               TabIndex        =   231
               Top             =   3030
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Bloqueo Regist.:"
               Height          =   315
               Index           =   89
               Left            =   -74880
               TabIndex        =   230
               Top             =   3690
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Comentarios Bloq.:"
               Height          =   465
               Index           =   91
               Left            =   -74880
               TabIndex        =   229
               Top             =   4350
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Aprob. Comité:"
               Height          =   315
               Index           =   85
               Left            =   -74880
               TabIndex        =   228
               Top             =   2370
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Firma Minuta:"
               Height          =   315
               Index           =   86
               Left            =   -74880
               TabIndex        =   227
               Top             =   2700
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inscrito en:"
               Height          =   315
               Index           =   90
               Left            =   -74880
               TabIndex        =   226
               Top             =   4020
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Tramitación de Póliza:"
               Height          =   435
               Index           =   83
               Left            =   -74880
               TabIndex        =   217
               Top             =   4320
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones de Evaluación:"
               Height          =   435
               Index           =   82
               Left            =   -74880
               TabIndex        =   213
               Top             =   3690
               Width           =   1605
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Tit.):"
               Height          =   285
               Index           =   70
               Left            =   -74880
               TabIndex        =   212
               Top             =   1410
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   72
               Left            =   -67470
               TabIndex        =   211
               Top             =   1410
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Préstamo:"
               Height          =   285
               Index           =   67
               Left            =   -74880
               TabIndex        =   210
               Top             =   420
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   73
               Left            =   -67470
               TabIndex        =   209
               Top             =   1080
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   68
               Left            =   -74880
               TabIndex        =   208
               Top             =   750
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   74
               Left            =   -67470
               TabIndex        =   207
               Top             =   750
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   69
               Left            =   -74880
               TabIndex        =   206
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza (Cyg.):"
               Height          =   285
               Index           =   71
               Left            =   -74880
               TabIndex        =   205
               Top             =   1740
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión:"
               Height          =   285
               Index           =   79
               Left            =   -67470
               TabIndex        =   204
               Top             =   3210
               Width           =   1095
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nro Póliza:"
               Height          =   285
               Index           =   78
               Left            =   -74880
               TabIndex        =   203
               Top             =   3210
               Width           =   1425
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Seguro de Vivienda:"
               Height          =   285
               Index           =   75
               Left            =   -74880
               TabIndex        =   202
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Factor/Importe:"
               Height          =   285
               Index           =   80
               Left            =   -67470
               TabIndex        =   201
               Top             =   2880
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   76
               Left            =   -74880
               TabIndex        =   200
               Top             =   2550
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Informe:"
               Height          =   285
               Index           =   81
               Left            =   -67470
               TabIndex        =   199
               Top             =   2550
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo de Aplicación:"
               Height          =   285
               Index           =   77
               Left            =   -74880
               TabIndex        =   198
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Emisión OT:"
               Height          =   315
               Index           =   57
               Left            =   -67470
               TabIndex        =   179
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Observaciones:"
               Height          =   315
               Index           =   66
               Left            =   -74880
               TabIndex        =   178
               Top             =   3810
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 1:"
               Height          =   285
               Index           =   62
               Left            =   -74880
               TabIndex        =   177
               Top             =   2220
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número Informe:"
               Height          =   285
               Index           =   59
               Left            =   -74880
               TabIndex        =   176
               Top             =   750
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "F. Evaluación:"
               Height          =   285
               Index           =   58
               Left            =   -67470
               TabIndex        =   175
               Top             =   750
               Width           =   1155
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Inmueble:"
               Height          =   285
               Index           =   61
               Left            =   -74880
               TabIndex        =   174
               Top             =   1890
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Empresa Peritaje:"
               Height          =   285
               Index           =   56
               Left            =   -74880
               TabIndex        =   173
               Top             =   420
               Width           =   1725
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nombre Perito:"
               Height          =   285
               Index           =   60
               Left            =   -74880
               TabIndex        =   172
               Top             =   1080
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estacionam. 2:"
               Height          =   285
               Index           =   63
               Left            =   -74880
               TabIndex        =   171
               Top             =   2550
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Depósito:"
               Height          =   285
               Index           =   64
               Left            =   -74880
               TabIndex        =   170
               Top             =   2880
               Width           =   1485
            End
            Begin VB.Line Line1 
               BorderWidth     =   2
               X1              =   -73230
               X2              =   -66510
               Y1              =   3270
               Y2              =   3270
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Totales:"
               Height          =   285
               Index           =   65
               Left            =   -74880
               TabIndex        =   169
               Top             =   3330
               Width           =   1485
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Relac. Cuota/Renta:"
               Height          =   315
               Index           =   55
               Left            =   -66510
               TabIndex        =   138
               Top             =   4200
               Width           =   1665
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. MPr.:"
               Height          =   315
               Index           =   54
               Left            =   -66510
               TabIndex        =   135
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. S/.:"
               Height          =   315
               Index           =   53
               Left            =   -70650
               TabIndex        =   133
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. US$:"
               Height          =   315
               Index           =   52
               Left            =   -74880
               TabIndex        =   131
               Top             =   1890
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio MPr.:"
               Height          =   315
               Index           =   51
               Left            =   -70650
               TabIndex        =   129
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Tipo Cambio US$:"
               Height          =   315
               Index           =   50
               Left            =   -74880
               TabIndex        =   127
               Top             =   4530
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Cónyuge S/.:"
               Height          =   315
               Index           =   49
               Left            =   -70650
               TabIndex        =   125
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "ILD Titular S/.:"
               Height          =   315
               Index           =   48
               Left            =   -74880
               TabIndex        =   123
               Top             =   4200
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Período de Gracia:"
               Height          =   315
               Index           =   47
               Left            =   -74880
               TabIndex        =   121
               Top             =   3870
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuotas Extraord.:"
               Height          =   315
               Index           =   46
               Left            =   -74880
               TabIndex        =   119
               Top             =   3540
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Plazo Aprobado:"
               Height          =   315
               Index           =   45
               Left            =   -74880
               TabIndex        =   117
               Top             =   3210
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Ultima Cuota:"
               Height          =   315
               Index           =   44
               Left            =   -74880
               TabIndex        =   115
               Top             =   2880
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Primera Cuota:"
               Height          =   315
               Index           =   43
               Left            =   -74880
               TabIndex        =   113
               Top             =   2550
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Cuota Fija:"
               Height          =   315
               Index           =   42
               Left            =   -74880
               TabIndex        =   111
               Top             =   2220
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. MPr.:"
               Height          =   315
               Index           =   41
               Left            =   -66510
               TabIndex        =   109
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. S/.:"
               Height          =   315
               Index           =   40
               Left            =   -70650
               TabIndex        =   107
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Solic. US$:"
               Height          =   315
               Index           =   39
               Left            =   -74880
               TabIndex        =   105
               Top             =   1410
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Aporte Propio US$:"
               Height          =   315
               Index           =   38
               Left            =   -74880
               TabIndex        =   103
               Top             =   1080
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Apellidos y Nombres:"
               Height          =   285
               Index           =   34
               Left            =   -74880
               TabIndex        =   101
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Documento Identidad:"
               Height          =   285
               Index           =   28
               Left            =   -74880
               TabIndex        =   99
               Top             =   420
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "E-Mail Personal:"
               Height          =   285
               Index           =   35
               Left            =   -68400
               TabIndex        =   96
               Top             =   1740
               Width           =   1245
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Celular:"
               Height          =   285
               Index           =   32
               Left            =   -68400
               TabIndex        =   95
               Top             =   1080
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Lugar de Nacimiento:"
               Height          =   285
               Index           =   31
               Left            =   -74880
               TabIndex        =   94
               Top             =   1740
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Profesión:"
               Height          =   285
               Index           =   30
               Left            =   -68400
               TabIndex        =   93
               Top             =   2070
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nivel de Estudios:"
               Height          =   285
               Index           =   29
               Left            =   -74880
               TabIndex        =   92
               Top             =   2070
               Width           =   1455
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "País de Nacimiento:"
               Height          =   285
               Index           =   27
               Left            =   -74880
               TabIndex        =   91
               Top             =   1410
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha de Nacimiento:"
               Height          =   285
               Index           =   26
               Left            =   -74880
               TabIndex        =   90
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Régimen Conyugal:"
               Height          =   285
               Index           =   5
               Left            =   6600
               TabIndex        =   72
               Top             =   1410
               Width           =   1395
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "E-Mail Personal:"
               Height          =   285
               Index           =   4
               Left            =   6600
               TabIndex        =   70
               Top             =   1080
               Width           =   1245
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   69
               Top             =   2070
               Width           =   1515
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Casa:"
               Height          =   285
               Index           =   3
               Left            =   6600
               TabIndex        =   68
               Top             =   750
               Width           =   1185
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Telf. Celular:"
               Height          =   285
               Index           =   1
               Left            =   6600
               TabIndex        =   67
               Top             =   420
               Width           =   1275
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Lugar de Nacimiento:"
               Height          =   285
               Index           =   19
               Left            =   120
               TabIndex        =   66
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Profesión:"
               Height          =   285
               Index           =   6
               Left            =   6600
               TabIndex        =   65
               Top             =   1740
               Width           =   1335
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Nivel de Estudios:"
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   64
               Top             =   1740
               Width           =   1455
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Estado Civil:"
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   63
               Top             =   1410
               Width           =   1365
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "País de Nacimiento:"
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   62
               Top             =   750
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha de Nacimiento:"
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   61
               Top             =   420
               Width           =   1695
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "V. Compra-Venta US$:"
               Height          =   315
               Index           =   37
               Left            =   -74880
               TabIndex        =   41
               Top             =   750
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda Préstamo:"
               Height          =   315
               Index           =   36
               Left            =   -74880
               TabIndex        =   40
               Top             =   420
               Width           =   1545
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            TabIndex        =   13
            Top             =   60
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Autorización de Desembolso"
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4920
            TabIndex        =   14
            Top             =   120
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "AteCli_frm_033.frx":016C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel34 
         Height          =   1725
         Left            =   30
         TabIndex        =   15
         Top             =   8700
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3043
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
         Begin VB.ComboBox cmb_BonoBP 
            Height          =   315
            Left            =   8850
            Style           =   2  'Dropdown List
            TabIndex        =   271
            Top             =   60
            Width           =   1305
         End
         Begin VB.ComboBox cmb_FueFin 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   270
            Top             =   60
            Width           =   3165
         End
         Begin VB.TextBox txt_Observ 
            Height          =   945
            Left            =   1800
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "AteCli_frm_033.frx":0476
            Top             =   720
            Width           =   10905
         End
         Begin EditLib.fpDateTime ipp_FecVal 
            Height          =   315
            Left            =   1800
            TabIndex        =   280
            Top             =   390
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
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
         Begin VB.Label Label9 
            Caption         =   "F. Valor Desemb.:"
            Height          =   285
            Left            =   60
            TabIndex        =   281
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label Label5 
            Caption         =   "Buen Pagador:"
            Height          =   315
            Left            =   7410
            TabIndex        =   273
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Financiamiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   272
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Comentarios:"
            Height          =   315
            Index           =   12
            Left            =   60
            TabIndex        =   16
            Top             =   720
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel36 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   2730
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   675
            Left            =   11310
            Picture         =   "AteCli_frm_033.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Rechaz 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_033.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Rechazar Solicitud"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "AteCli_frm_033.frx":0BC6
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "AteCli_frm_033.frx":1008
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10560
            Picture         =   "AteCli_frm_033.frx":1312
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   675
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   23
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label35 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   22
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label31 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   21
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   24
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1931
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   25
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   8850
            TabIndex        =   26
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_EjeVta 
            Height          =   315
            Left            =   8850
            TabIndex        =   27
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   28
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   29
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_IniEva 
            Height          =   315
            Left            =   8850
            TabIndex        =   30
            Top             =   390
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "31/12/2004"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Index           =   15
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "F. Ingreso Solic.:"
            Height          =   315
            Index           =   18
            Left            =   7410
            TabIndex        =   35
            Top             =   60
            Width           =   1185
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Ejecutivo Ventas:"
            Height          =   315
            Index           =   16
            Left            =   7410
            TabIndex        =   34
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Modalidad:"
            Height          =   315
            Index           =   13
            Left            =   60
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Producto:"
            Height          =   315
            Index           =   14
            Left            =   60
            TabIndex        =   32
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "F. Inicio Evaluac.:"
            Height          =   315
            Index           =   17
            Left            =   7410
            TabIndex        =   31
            Top             =   390
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_AutDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_IniEva     As String
Dim l_str_Aprueb     As String
Dim l_str_Rechaz     As String
Dim l_arr_FueFin()   As moddat_tpo_Genera

Private Sub cmb_BonoBP_Click()
   Call gs_SetFocus(ipp_FecVal)
End Sub

Private Sub cmb_BonoBP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BonoBP_Click
   End If
End Sub

Private Sub cmb_FueFin_Click()
   Call gs_SetFocus(cmb_BonoBP)
End Sub

Private Sub cmb_FueFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FueFin_Click
   End If
End Sub

Private Sub cmd_Aprueb_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   If cmb_FueFin.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Fuente de Financiamiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FueFin)
      Exit Sub
   End If
   
   If cmb_BonoBP.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente tiene beneficio de Bono Buen Pagador.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BonoBP)
      Exit Sub
   End If
   
   If CDate(ipp_FecVal.Text) < Date Then
      MsgBox "La Fecha de Valor no puede ser menor a la Fecha de Hoy.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecVal)
      Exit Sub
   End If
   
   If Len(Trim(txt_Observ.Text)) = 0 Then
      MsgBox "Debe ingresar las observaciones para el Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Observ)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de aprobar esta instancia de Evaluación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Call moddat_gs_FecSis
   r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
   
   'Grabando Autorización de Desembolso
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_AUTDES ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_FueFin(cmb_FueFin.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_BonoBP.ItemData(cmb_BonoBP.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecVal.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
            
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_AUTDES. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   'Actualizando en Instancia
   If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_AutDes, r_int_DiaTra, 1, 1) Then
      Exit Sub
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_AutDes, 12, 0, txt_Observ.Text, 0, 0) Then
      Exit Sub
   End If
   
   'Inserta Nueva Instancia de Evaluación
   If Not moddat_gf_Inserta_Seguim(moddat_g_str_NumSol, modatecli_g_con_Desemb) Then
      Exit Sub
   End If

   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_Desemb, 11, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Actualizando en Tabla de Créditos
   If Not modatecli_gf_ActIns_SolMae(moddat_g_str_NumSol, modatecli_g_con_Desemb) Then
      Exit Sub
   End If
   
   r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   r_str_Cadena = r_str_Cadena & Chr(13)

   modgen_g_str_Mail_Asunto = "AUTORIZACION DE DESEMBOLSO APROBADA (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   modgen_g_str_Mail_Mensaj = r_str_Cadena
   
   frm_EnvMai_01.Show 1
   
   MsgBox "Se aprobo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
      g_str_Parame = g_str_Parame & "SOLMAE_ENVCRE = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   pnl_EjeVta.Caption = moddat_g_str_EjeVta
   pnl_FecIng.Caption = moddat_g_str_FecIng

   'Validación que se encuentre en Instancia
   If moddat_g_int_InsAct <> modatecli_g_con_AutDes Then
      MsgBox "No se encuentra en Instancia de Autorización de Desembolso.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   l_str_IniEva = ""
   l_str_Aprueb = ""
   l_str_Rechaz = ""

   'Obteniendo Información del Seguimiento
   Call fs_Buscar_SegDet
   
   If Len(Trim(l_str_Aprueb)) > 0 Then
      MsgBox "El cliente ya ha sido aprobado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Len(Trim(l_str_Rechaz)) > 0 Then
      MsgBox "El cliente ya ha sido rechazado en esta instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_Buscar_InfSol
   Call fs_ActivaItem(True)
   
   Call gs_SetFocus(cmb_FueFin)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Rechaz_Click()
   Dim r_int_DiaTra     As Integer
   Dim r_str_CodIns     As String
   Dim r_str_Cadena     As String
   
   moddat_g_int_InsAct = modatecli_g_con_EvaTas
   moddat_g_int_MotRec = 0
   moddat_g_str_Observ = ""
   
   frm_Rechaz_01.Show 1
   
   If moddat_g_int_MotRec > 0 Then
      Call moddat_gs_FecSis
      r_int_DiaTra = CInt(CDate(moddat_g_str_FecSis) - CDate(l_str_IniEva))
      
      'Actualizando en Instancia
      If Not moddat_gf_Modifica_Seguim(moddat_g_str_NumSol, modatecli_g_con_AutDes, r_int_DiaTra, 2, 1) Then
         Exit Sub
      End If
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, modatecli_g_con_AutDes, 13, 0, moddat_g_str_Observ, 0, moddat_g_int_MotRec) Then
         Exit Sub
      End If
      
      'Actualizando Rechazo en Tabla de Créditos
      If Not modatecli_gf_Rechaz_SolMae(moddat_g_str_NumSol, 1, moddat_g_int_MotRec) Then
         Exit Sub
      End If
   
      r_str_Cadena = r_str_Cadena & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      r_str_Cadena = r_str_Cadena & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      r_str_Cadena = r_str_Cadena & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      r_str_Cadena = r_str_Cadena & Chr(13)
   
      modgen_g_str_Mail_Asunto = "RECHAZO EN AUTORIZACION DE DESEMBOLSO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      modgen_g_str_Mail_Mensaj = r_str_Cadena
      
      frm_EnvMai_01.Show 1
      
      MsgBox "Se rechazo la Solicitud en esta Instancia de Evaluación.", vbInformation, modgen_g_str_NomPlt
      
      Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub fs_Inicia()
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   
   Call moddat_gs_Carga_LisIte(cmb_FueFin, l_arr_FueFin(), 1, "502")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BonoBP, 1, "214")
   
   grd_Tit_ActPri.ColWidth(0) = 2130
   grd_Tit_ActPri.ColWidth(1) = 8200
   grd_Tit_ActPri.ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_ActPri.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Tit_ActSec.ColWidth(0) = 2130
   grd_Tit_ActSec.ColWidth(1) = 8200
   grd_Tit_ActSec.ColAlignment(0) = flexAlignLeftCenter
   grd_Tit_ActSec.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_ActPri.ColWidth(0) = 2130
   grd_Cyg_ActPri.ColWidth(1) = 8200
   grd_Cyg_ActPri.ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_ActPri.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Cyg_ActSec.ColWidth(0) = 2130
   grd_Cyg_ActSec.ColWidth(1) = 8200
   grd_Cyg_ActSec.ColAlignment(0) = flexAlignLeftCenter
   grd_Cyg_ActSec.ColAlignment(1) = flexAlignLeftCenter
   
End Sub

Private Sub ipp_FecVal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Limpia()
   Call fs_ActivaItem(False)
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   pnl_NumSol.Caption = ""
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_EjeVta.Caption = ""
   pnl_FecIng.Caption = ""
   pnl_IniEva.Caption = ""
   
   tab_Princi.Tab = 0
   tab_DatCli.Tab = 0
   tab_DatCyg.Tab = 0
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_LimpiaItem()
   'Datos del Cliente
   pnl_Tit_FecNac.Caption = ""
   pnl_Tit_Paises.Caption = ""
   pnl_Tit_LugNac.Caption = ""
   pnl_Tit_EstCiv.Caption = ""
   pnl_Tit_NivEst.Caption = ""
   pnl_Tit_Direcc.Caption = ""
   pnl_Tit_Celula.Caption = ""
   pnl_Tit_Telefo.Caption = ""
   pnl_Tit_DirEle.Caption = ""
   pnl_Tit_RegCyg.Caption = ""
   pnl_Tit_Profes.Caption = ""

   pnl_Tit_OcuPri.Caption = ""
   Call gs_LimpiaGrid(grd_Tit_ActPri)
   
   pnl_Tit_OcuSec.Caption = ""
   Call gs_LimpiaGrid(grd_Tit_ActSec)

   'Datos del Cónyuge
   pnl_Cyg_DocIde.Caption = ""
   pnl_Cyg_ApeNom.Caption = ""
   pnl_Cyg_FecNac.Caption = ""
   pnl_Cyg_Paises.Caption = ""
   pnl_Cyg_LugNac.Caption = ""
   pnl_Cyg_NivEst.Caption = ""
   pnl_Cyg_Celula.Caption = ""
   pnl_Cyg_DirEle.Caption = ""
   pnl_Cyg_Profes.Caption = ""
   
   pnl_Cyg_OcuPri.Caption = ""
   Call gs_LimpiaGrid(grd_Cyg_ActPri)
   
   pnl_Cyg_OcuSec.Caption = ""
   Call gs_LimpiaGrid(grd_Cyg_ActSec)
   
   'Datos de Crédito
   pnl_Cre_TipMon.Caption = ""
   pnl_Cre_TasInt.Caption = "0.00 "
   pnl_Cre_ComVta.Caption = "0.00 "
   pnl_Cre_ApoPro.Caption = "0.00 "
   pnl_Cre_MonSol_Dol.Caption = "0.00 "
   pnl_Cre_MonSol_Sol.Caption = "0.00 "
   pnl_Cre_MonSol_MPr.Caption = "0.00 "
   pnl_Cre_MonApr_Dol.Caption = "0.00 "
   pnl_Cre_MonApr_Sol.Caption = "0.00 "
   pnl_Cre_MonApr_MPr.Caption = "0.00 "
   pnl_Cre_CuoFij.Caption = "0.00 "
   pnl_Cre_CuoIni.Caption = "0.00 "
   pnl_Cre_CuoFin.Caption = "0.00 "
   pnl_Cre_PlaApr.Caption = "0 "
   pnl_Cre_CuoExt.Caption = ""
   pnl_Cre_PerGra.Caption = "0 "
   pnl_Cre_ILDTit.Caption = "0.00 "
   pnl_Cre_ILDCyg.Caption = "0.00 "
   pnl_Cre_CuoRen.Caption = "0.00 "
   pnl_Cre_TCaDol.Caption = "0.000000000 "
   pnl_Cre_TCaMPr.Caption = "0.000000000 "

   'Datos de Tasación
   pnl_Tas_EmpPer.Caption = ""
   pnl_Tas_NumInf.Caption = ""
   pnl_Tas_FecEmi.Caption = ""
   pnl_Tas_FecEva.Caption = ""
   pnl_Tas_NomPer.Caption = ""
   pnl_Tas_ValCom.Caption = "0.00 "
   pnl_Tas_ValRea.Caption = "0.00 "
   pnl_Tas_AreTer.Caption = "0.00 "
   pnl_Tas_AreCon.Caption = "0.00 "
   pnl_Tas_VCoEs1.Caption = "0.00 "
   pnl_Tas_VReEs1.Caption = "0.00 "
   pnl_Tas_ATeEs1.Caption = "0.00 "
   pnl_Tas_ACoEs1.Caption = "0.00 "
   pnl_Tas_VCoEs2.Caption = "0.00 "
   pnl_Tas_VReEs2.Caption = "0.00 "
   pnl_Tas_ATeEs2.Caption = "0.00 "
   pnl_Tas_ACoEs2.Caption = "0.00 "
   pnl_Tas_VCoDep.Caption = "0.00 "
   pnl_Tas_VReDep.Caption = "0.00 "
   pnl_Tas_ATeDep.Caption = "0.00 "
   pnl_Tas_ACoDep.Caption = "0.00 "
   pnl_Tas_TotVCo.Caption = "0.00 "
   pnl_Tas_TotVRe.Caption = "0.00 "
   pnl_Tas_TotATe.Caption = "0.00 "
   pnl_Tas_TotACo.Caption = "0.00 "
   txt_Tas_Observ.Text = ""

   'Datos de Seguros
   pnl_Seg_SegPre.Caption = ""
   pnl_Seg_SegViv.Caption = ""
   pnl_Seg_InfPre.Caption = ""
   pnl_Seg_EvaPre.Caption = ""
   pnl_Seg_AplPre.Caption = ""
   pnl_Seg_FoiPre.Caption = ""
   pnl_Seg_PolTit.Caption = ""
   pnl_Seg_EmiTit.Caption = ""
   pnl_Seg_PolCyg.Caption = ""
   pnl_Seg_InfViv.Caption = ""
   pnl_Seg_EvaViv.Caption = ""
   pnl_Seg_AplViv.Caption = ""
   pnl_Seg_FoiViv.Caption = ""
   pnl_Seg_PolViv.Caption = ""
   pnl_Seg_EmiViv.Caption = ""
   txt_Seg_ObsEva.Text = ""
   txt_Seg_ObsPol.Text = ""
   
   'Datos de Legal
   txt_Leg_InfLeg.Text = ""
   pnl_Leg_AprCom.Caption = ""
   pnl_Leg_FirCon.Caption = ""
   pnl_Leg_RepLeg.Caption = ""
   pnl_Leg_Notari.Caption = ""
   pnl_Leg_FecBlq.Caption = ""
   pnl_Leg_DocReg.Caption = ""
   txt_Leg_ObsBlq.Text = ""

   'Datos de COFIDE
   pnl_Cof_NumCar.Caption = ""
   pnl_Cof_FecEmi.Caption = ""
   pnl_Cof_FecVal.Caption = ""
   pnl_Cof_NomBan.Caption = ""
   pnl_Cof_NumCta.Caption = ""
   pnl_Cof_NumOpe.Caption = ""
   pnl_Cof_TipMon.Caption = ""
   pnl_Cof_Import.Caption = "0.00 "
   pnl_Cof_TasInt.Caption = "0.00 "
   pnl_Cof_TasCom.Caption = "0.00 "
   
   'Datos de Inmueble
   pnl_Inm_Direcc.Caption = ""
   pnl_Inm_TipPro.Caption = ""
   pnl_Inm_JurEmp.Caption = ""
   pnl_Inm_JurRep.Caption = ""
   pnl_Inm_JurDir.Caption = ""
   pnl_Inm_NatTit.Caption = ""
   pnl_Inm_NatCyg.Caption = ""
   
   'Autorización de Desembolso
   cmb_FueFin.ListIndex = -1
   cmb_BonoBP.ListIndex = -1
   ipp_FecVal.Text = Format(Date, "dd/mm/yyyy")
   txt_Observ.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   tab_Princi.Enabled = Not p_Habilita
   cmd_Aprueb.Enabled = Not p_Habilita
   cmd_Rechaz.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   txt_Observ.Enabled = p_Habilita
   cmb_FueFin.Enabled = p_Habilita
   ipp_FecVal.Enabled = p_Habilita
   cmb_BonoBP.Enabled = p_Habilita
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!SOLMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!SOLMAE_CODMOD & "")
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!SOLMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Ventas
   moddat_g_str_CodEje = Trim(g_rst_Princi!SOLMAE_EJEVTA)
   moddat_g_str_EjeVta = moddat_gf_Buscar_NomEje(moddat_g_str_CodEje)

   'Instancia Actual
   moddat_g_int_InsAct = g_rst_Princi!SOLMAE_CODINS

   'Moneda
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))

   'Fecha de Ingreso
   moddat_g_str_FecIng = Right(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 2) & "/" & Mid(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 5, 2) & "/" & Left(Format(g_rst_Princi!SOLMAE_FECSOL, "00000000"), 4)
   
   'Información de Seguros
   pnl_Seg_SegPre.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES)) & " / " & Trim(moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG))
   pnl_Seg_SegViv.Caption = Trim(moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGVIV))
End Sub

Private Sub fs_Buscar_SegDet()
   Dim r_str_FecOcu  As String
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(modatecli_g_con_AutDes) & " "
   g_str_Parame = g_str_Parame & "ORDER BY SEGFECCRE DESC, SEGHORCRE DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_FecOcu = Right(CStr(g_rst_Princi!SEGDET_FECOCU), 2) & "/" & Mid(CStr(g_rst_Princi!SEGDET_FECOCU), 5, 2) & "/" & Left(CStr(g_rst_Princi!SEGDET_FECOCU), 4)
      
      Select Case g_rst_Princi!SEGDET_CODOCU
         Case 11:    l_str_IniEva = r_str_FecOcu
         Case 12:    l_str_Aprueb = r_str_FecOcu
         Case 13:    l_str_Rechaz = r_str_FecOcu
      End Select
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If Len(Trim(l_str_IniEva)) > 0 Then
      pnl_IniEva.Caption = l_str_IniEva
   End If
End Sub

Private Sub fs_Buscar_InfSol()
   Call fs_Buscar_DatCli
   Call fs_Buscar_DatCyg
   Call fs_Buscar_DatCre
   Call fs_Buscar_DatTas
   Call fs_Buscar_DatSeg
   Call fs_Buscar_DatLeg
   
   If moddat_g_str_CodPrd = "001" Then
      Call fs_Buscar_DatCof
   End If
   
   Call fs_Buscar_DatInm
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmd_Aprueb.Enabled Then
         Call gs_SetFocus(cmd_Aprueb)
      Else
         Call gs_SetFocus(cmd_Rechaz)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Buscar_DatCli()
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   pnl_Tit_FecNac.Caption = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
   pnl_Tit_Celula.Caption = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
   pnl_Tit_Telefo.Caption = Trim(g_rst_Princi!DatGen_Telefo & "")
   pnl_Tit_DirEle.Caption = Trim(g_rst_Princi!DatGen_DirEle & "")

   pnl_Tit_EstCiv.Caption = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DatGen_EstCiv))
   
   If g_rst_Princi!DatGen_RegCyg > 0 Then
      pnl_Tit_RegCyg.Caption = moddat_gf_Consulta_ParDes("206", CStr(g_rst_Princi!DatGen_RegCyg))
   End If
   
   pnl_Tit_NivEst.Caption = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

   'País de Nacimiento
   pnl_Tit_Paises.Caption = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))

   'Profesión
   pnl_Tit_Profes.Caption = moddat_gf_Consulta_ParDes("501", Trim(g_rst_Princi!DatGen_Profes))

   r_str_Depart = ""
   r_str_Provin = ""
   r_str_Distri = ""
   
   If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
      'Departamento
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000")
      
      'Provincia
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00")
      
      'Distrito
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      
      pnl_Tit_LugNac.Caption = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If

   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon))

   pnl_Tit_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero)
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_IntDpt))) > 0 Then
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!DatGen_NomZon))) > 0 Then
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!DatGen_NomZon) & Chr(13) & Chr(10)
   Else
      pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
   
   pnl_Tit_Direcc.Caption = pnl_Tit_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   If g_rst_Princi!DatGen_CygTDo > 0 Then
      moddat_g_int_CygTDo = g_rst_Princi!DatGen_CygTDo
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DatGen_CygNDo)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1, grd_Tit_ActPri, pnl_Tit_OcuPri)
   Call fs_Buscar_ActEco(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2, grd_Tit_ActSec, pnl_Tit_OcuSec)
End Sub

Private Sub fs_Buscar_DatCyg()
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String

   If moddat_g_int_CygTDo = 0 Then
      Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   g_rst_Princi.MoveFirst

   pnl_Cyg_DocIde.Caption = CStr(g_rst_Princi!DatGen_TipDoc) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   pnl_Cyg_ApeNom.Caption = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)

   pnl_Cyg_FecNac.Caption = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
   pnl_Cyg_Celula.Caption = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
   pnl_Cyg_DirEle.Caption = Trim(g_rst_Princi!DatGen_DirEle & "")

   pnl_Cyg_NivEst.Caption = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

   'País de Nacimiento
   pnl_Cyg_Paises.Caption = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))

   'Profesión
   pnl_Cyg_Profes.Caption = moddat_gf_Consulta_ParDes("501", Trim(g_rst_Princi!DatGen_Profes))

   r_str_Depart = ""
   r_str_Provin = ""
   r_str_Distri = ""
   
   If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
      'Departamento
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000")
      
      'Provincia
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00")
      
      'Distrito
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      
      pnl_Cyg_LugNac.Caption = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1, grd_Cyg_ActPri, pnl_Cyg_OcuPri)
   Call fs_Buscar_ActEco(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2, grd_Cyg_ActSec, pnl_Cyg_OcuSec)
End Sub

Private Sub fs_Buscar_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, p_Listad As MSFlexGrid, p_NomOcu As SSPanel)
   Dim r_str_Depart     As String
   Dim r_str_Provin     As String
   Dim r_str_Distri     As String
   Dim r_str_TipVia     As String
   Dim r_str_TipZon     As String
   Dim r_str_TipDoc     As String
   Dim l_rst_Genera     As ADODB.Recordset
   
   Call gs_LimpiaGrid(p_Listad)
   p_NomOcu.Caption = ""
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Ocupación
   p_NomOcu.Caption = moddat_gf_Consulta_ParDes("008", CStr(g_rst_Princi!ActEco_CodAct))
   
   Select Case g_rst_Princi!ActEco_CodAct
      Case 11, 31, 41
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_NumDoc) & "' "
      
         If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
            Exit Sub
         End If
   
         l_rst_Genera.MoveFirst
         
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
      
         'Razón Social
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Razón Social"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
      
         'Nombre Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Nombre Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
      
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
      
         If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
         End If
      
         'Dirección
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Dirección Empresa"
      
         p_Listad.Col = 1
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))

         p_Listad.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")

         If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
         End If

         If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
         Else
            p_Listad.Text = p_Listad.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
   
         p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
         'Teléfono
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Teléfono(s) Empresa"
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
         If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         'Sucursal
         If Len(Trim(g_rst_Princi!ActEco_Sucurs & "")) > 0 Then
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Sucursal"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ACTECO_DEP_SUCURS & "")
            
            'Dirección Sucursal
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Dirección Sucursal"
         
            p_Listad.Col = 1
            
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

            p_Listad.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
            Else
               p_Listad.Text = p_Listad.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
      
            p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfono Sucursal
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Teléfono(s) Sucursal"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
            
            If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
            End If
         End If
         
         If g_rst_Princi!ActEco_CodAct = 11 Then
            'Teléfono y Anexo RR.HH
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Teléfono RR.HH"
         
            p_Listad.Col = 1
            
            If Len(Trim(l_rst_Genera!DATGEN_TELERH & "")) = 0 Then
               p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
            Else
               p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELERH & "")
            End If
            
            If Len(Trim(l_rst_Genera!DATGEN_ANEXRH & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_ANEXRH & "")
            End If
         
            'Cargo
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Cargo"
         
            p_Listad.Col = 1
            If Len(Trim(g_rst_Princi!ActEco_Dep_CargoN & "")) > 0 Then
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_CargoN)
            Else
               p_Listad.Text = moddat_gf_Consulta_ParDes("503", Trim(g_rst_Princi!ActEco_Dep_CargoC))
            End If
         
            'Area
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Area"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_NomAre)
            
            'Número Anexo
            If Len(Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Anexo"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx)
            End If
            
            'Teléfono Directo
            If Len(Trim(g_rst_Princi!ActEco_Dep_TelDir & "")) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Teléfono Directo"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir)
            End If
         
            'Celular Laboral
            If Len(Trim(g_rst_Princi!ActEco_Dep_Celula)) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "Celular Laboral"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_Celula)
            End If
         
            'E-mail
            If Len(Trim(g_rst_Princi!ActEco_Dep_DirEle)) > 0 Then
               p_Listad.Rows = p_Listad.Rows + 1
               p_Listad.Row = p_Listad.Rows - 1
               
               p_Listad.Col = 0
               p_Listad.Text = "E-mail"
            
               p_Listad.Col = 1
               p_Listad.Text = Trim(g_rst_Princi!ActEco_Dep_DirEle)
            End If
         End If
         
         l_rst_Genera.Close
         Set l_rst_Genera = Nothing
         
         Call gs_UbiIniGrid(p_Listad)
         
      Case 21
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Dirección Tributaria
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Dirección Tributaria"
      
         p_Listad.Col = 1
         
         r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_TipVia))
         r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_TipZon))

         p_Listad.Text = r_str_TipVia & " " & Trim(g_rst_Princi!ActEco_NomVia & "") & " " & Trim(g_rst_Princi!ActEco_Numero & "")

         If Len(Trim(Trim(g_rst_Princi!ActEco_IntDpt & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " (" & Trim(g_rst_Princi!ActEco_IntDpt) & ")"
         End If

         If Len(Trim(Trim(g_rst_Princi!ActEco_NomZon & ""))) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!ActEco_NomZon) & " / "
         Else
            p_Listad.Text = p_Listad.Text & " / "
         End If
         
         r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 2) & "0000")
         r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ubigeo, 4) & "00")
         r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ubigeo))
   
         p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
      
         'Teléfono
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Teléfono(s) "
      
         p_Listad.Col = 1
         p_Listad.Text = Trim(g_rst_Princi!ActEco_Telef1 & "")
         
         If Len(Trim(g_rst_Princi!ActEco_Telef2 & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & Trim(g_rst_Princi!ActEco_Telef2 & "")
         End If
         
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         'Contrato de Locación de Servicios
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Contrato Locación "
         
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
         
         If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
            g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TDoEmp) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & Trim(g_rst_Princi!ActEco_Ind_NDoEmp) & "' "
      
            If Not gf_EjecutaSQL(g_str_Parame, l_rst_Genera, 3) Then
               Exit Sub
            End If
   
            l_rst_Genera.MoveFirst
         
            'Documento de Identidad
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Documento Ident. Empresa"
      
            p_Listad.Col = 1
            p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(l_rst_Genera!DatGen_EMPTDO)) & " - " & Trim(l_rst_Genera!DatGen_EMPNDO)
      
            'Razón Social
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Razón Social Empresa"
      
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_RAZSOC)
         
            'Nombre Comercial
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Nombre Comercial Empresa"
         
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_NOMCOM)
         
            'Giro Comercial
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Giro Comercial Empresa"
         
            p_Listad.Col = 1
            p_Listad.Text = moddat_gf_Busca_GirCom(Trim(l_rst_Genera!DATGEN_GCOMCO))
         
            If Len(Trim(l_rst_Genera!DATGEN_GCOMNO & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & Trim(l_rst_Genera!DATGEN_GCOMNO)
            End If
         
            'Dirección
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            p_Listad.Col = 0
            p_Listad.Text = "Dirección Empresa"
         
            p_Listad.Col = 1
            r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Genera!DatGen_TipVia))
            r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Genera!DatGen_TipZon))
   
            p_Listad.Text = r_str_TipVia & " " & Trim(l_rst_Genera!DatGen_NomVia & "") & " " & Trim(l_rst_Genera!DatGen_Numero & "")
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_IntDpt & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " (" & Trim(l_rst_Genera!DatGen_IntDpt) & ")"
            End If
   
            If Len(Trim(Trim(l_rst_Genera!DatGen_NomZon & ""))) > 0 Then
               p_Listad.Text = p_Listad.Text & " - " & r_str_TipZon & " " & Trim(l_rst_Genera!DatGen_NomZon) & " / "
            Else
               p_Listad.Text = p_Listad.Text & " / "
            End If
            
            r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 2) & "0000")
            r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Genera!DatGen_Ubigeo, 4) & "00")
            r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Genera!DatGen_Ubigeo))
      
            p_Listad.Text = p_Listad.Text & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
         
            'Teléfonos
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
         
            p_Listad.Col = 0
            p_Listad.Text = "Teléfonos Empresa"
            
            p_Listad.Col = 1
            p_Listad.Text = Trim(l_rst_Genera!DATGEN_TELEF1 & "")
         
            If Len(Trim(l_rst_Genera!DATGEN_TELEF2 & "")) > 0 Then
               p_Listad.Text = p_Listad.Text & Trim(l_rst_Genera!DATGEN_TELEF2 & "")
            End If
         End If
         
         Call gs_UbiIniGrid(p_Listad)
         
      Case 51
         'Documento de Identidad
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Documento de Identidad"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!ActEco_TipDoc)) & " - " & Trim(g_rst_Princi!ActEco_NumDoc)
         
         'Giro Comercial
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 0
         p_Listad.Text = "Giro Comercial"
      
         p_Listad.Col = 1
         p_Listad.Text = moddat_gf_Busca_GirCom(Trim(g_rst_Princi!ActEco_GiroCd))
      
         If Len(Trim(g_rst_Princi!ActEco_GiroNm & "")) > 0 Then
            p_Listad.Text = p_Listad.Text & " - " & Trim(g_rst_Princi!ActEco_GiroNm)
         End If
         
         Call gs_UbiIniGrid(p_Listad)
   End Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatCre()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst

   pnl_Cre_TipMon.Caption = moddat_gf_Consulta_ParDes("204", g_rst_Princi!SOLMAE_TIPMON)
   pnl_Cre_TasInt.Caption = Format(g_rst_Princi!SOLMAE_TASINT, "##0.00") & " "
   pnl_Cre_ComVta.Caption = Format(g_rst_Princi!SOLMAE_COMVTA, "###,###,##0.00") & " "
   pnl_Cre_ApoPro.Caption = Format(g_rst_Princi!SOLMAE_APOPRO, "###,###,##0.00") & " "
   pnl_Cre_MonSol_Dol.Caption = Format(g_rst_Princi!SOLMAE_MTOSOL, "###,###,##0.00") & " "
   pnl_Cre_MonSol_Sol.Caption = Format(g_rst_Princi!SOLMAE_PRESOL, "###,###,##0.00") & " "
   pnl_Cre_MonSol_MPr.Caption = Format(g_rst_Princi!SOLMAE_PREMPR, "###,###,##0.00") & " "
   pnl_Cre_MonApr_Dol.Caption = Format(g_rst_Princi!SOLMAE_APRDOL, "###,###,##0.00") & " "
   pnl_Cre_MonApr_Sol.Caption = Format(g_rst_Princi!SOLMAE_APRSOL, "###,###,##0.00") & " "
   pnl_Cre_MonApr_MPr.Caption = Format(g_rst_Princi!SOLMAE_APRMPR, "###,###,##0.00") & " "
   pnl_Cre_CuoFij.Caption = Format(g_rst_Princi!SOLMAE_APRCUO, "###,###,##0.00") & " "
   pnl_Cre_CuoIni.Caption = Format(g_rst_Princi!SOLMAE_APRCIN, "###,###,##0.00") & " "
   pnl_Cre_CuoFin.Caption = Format(g_rst_Princi!SOLMAE_APRCFN, "###,###,##0.00") & " "
   pnl_Cre_PlaApr.Caption = Format(g_rst_Princi!SOLMAE_APRPLA, "##0") & " "
   pnl_Cre_PerGra.Caption = Format(g_rst_Princi!SOLMAE_APRPGR, "##0") & " "
   pnl_Cre_CuoExt.Caption = moddat_gf_Consulta_ParDes("223", g_rst_Princi!SOLMAE_CUOANO)
   pnl_Cre_ILDTit.Caption = Format(g_rst_Princi!SOLMAE_APRIN1 + g_rst_Princi!SOLMAE_APRIN2, "###,###,##0.00") & " "
   pnl_Cre_ILDCyg.Caption = Format(g_rst_Princi!SOLMAE_APRIN3 + g_rst_Princi!SOLMAE_APRIN4, "###,###,##0.00") & " "
   pnl_Cre_CuoRen.Caption = Format(g_rst_Princi!SOLMAE_APRRCR, "##0.00") & " "
   pnl_Cre_TCaDol.Caption = Format(g_rst_Princi!SOLMAE_APRTCD, "###,##0.000000") & " "
   pnl_Cre_TCaMPr.Caption = Format(g_rst_Princi!SOLMAE_APRTCM, "###,##0.000000") & " "
   
   pnl_Cof_TasInt.Caption = Format(g_rst_Princi!SOLMAE_INTCOF, "##0.00") & " "
   pnl_Cof_TasCom.Caption = Format(g_rst_Princi!SOLMAE_COMCOF, "##0.00") & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatTas()
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   'Empresa de Peritaje
   pnl_Tas_EmpPer.Caption = moddat_gf_Consulta_ParDes("507", Format(g_rst_Princi!EVATAS_CODEMP, "000000"))

   pnl_Tas_NumInf.Caption = Trim(g_rst_Princi!EVATAS_NUMINF)
   pnl_Tas_FecEmi.Caption = Right(CStr(g_rst_Princi!EVATAS_FECEMI), 2) & "/" & Mid(CStr(g_rst_Princi!EVATAS_FECEMI), 5, 2) & "/" & Left(CStr(g_rst_Princi!EVATAS_FECEMI), 4)
   pnl_Tas_FecEva.Caption = Right(CStr(g_rst_Princi!EVATAS_FECEVA), 2) & "/" & Mid(CStr(g_rst_Princi!EVATAS_FECEVA), 5, 2) & "/" & Left(CStr(g_rst_Princi!EVATAS_FECEVA), 4)
   pnl_Tas_NomPer.Caption = Trim(g_rst_Princi!EVATAS_NOMPER)
   
   pnl_Tas_ValCom.Caption = Format(g_rst_Princi!EVATAS_VALCOM, "###,###,##0.00") & " "
   pnl_Tas_ValRea.Caption = Format(g_rst_Princi!EVATAS_VALFAB, "###,###,##0.00") & " "
   pnl_Tas_AreTer.Caption = Format(g_rst_Princi!EVATAS_ARETER, "###,###,##0.00") & " "
   pnl_Tas_AreCon.Caption = Format(g_rst_Princi!EVATAS_ARECON, "###,###,##0.00") & " "
   
   pnl_Tas_VCoEs1.Caption = Format(g_rst_Princi!EVATAS_VCOES1, "###,###,##0.00") & " "
   pnl_Tas_VReEs1.Caption = Format(g_rst_Princi!EVATAS_VREES1, "###,###,##0.00") & " "
   pnl_Tas_ATeEs1.Caption = Format(g_rst_Princi!EVATAS_ATEES1, "###,###,##0.00") & " "
   pnl_Tas_ACoEs1.Caption = Format(g_rst_Princi!EVATAS_ACOES1, "###,###,##0.00") & " "
   
   pnl_Tas_VCoEs2.Caption = Format(g_rst_Princi!EVATAS_VCOES2, "###,###,##0.00") & " "
   pnl_Tas_VReEs2.Caption = Format(g_rst_Princi!EVATAS_VREES2, "###,###,##0.00") & " "
   pnl_Tas_ATeEs2.Caption = Format(g_rst_Princi!EVATAS_ATEES2, "###,###,##0.00") & " "
   pnl_Tas_ACoEs2.Caption = Format(g_rst_Princi!EVATAS_ACOES2, "###,###,##0.00") & " "
   
   pnl_Tas_VCoDep.Caption = Format(g_rst_Princi!EVATAS_VCODEP, "###,###,##0.00") & " "
   pnl_Tas_VReDep.Caption = Format(g_rst_Princi!EVATAS_VREDEP, "###,###,##0.00") & " "
   pnl_Tas_ATeDep.Caption = Format(g_rst_Princi!EVATAS_ATEDEP, "###,###,##0.00") & " "
   pnl_Tas_ACoDep.Caption = Format(g_rst_Princi!EVATAS_ACODEP, "###,###,##0.00") & " "
   
   pnl_Tas_TotVCo.Caption = Format(g_rst_Princi!EVATAS_VALCOM + g_rst_Princi!EVATAS_VCOES1 + g_rst_Princi!EVATAS_VCOES2 + g_rst_Princi!EVATAS_VCODEP, "###,###,##0.00") & " "
   pnl_Tas_TotVRe.Caption = Format(g_rst_Princi!EVATAS_VALFAB + g_rst_Princi!EVATAS_VREES1 + g_rst_Princi!EVATAS_VREES2 + g_rst_Princi!EVATAS_VREDEP, "###,###,##0.00") & " "
   pnl_Tas_TotATe.Caption = Format(g_rst_Princi!EVATAS_ARETER + g_rst_Princi!EVATAS_ATEES1 + g_rst_Princi!EVATAS_ATEES2 + g_rst_Princi!EVATAS_ATEDEP, "###,###,##0.00") & " "
   pnl_Tas_TotACo.Caption = Format(g_rst_Princi!EVATAS_ARECON + g_rst_Princi!EVATAS_ACOES1 + g_rst_Princi!EVATAS_ACOES2 + g_rst_Princi!EVATAS_ACODEP, "###,###,##0.00") & " "
   
   txt_Tas_Observ.Text = Trim(g_rst_Princi!EVATAS_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatLeg()
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Leg_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLEG)
   
   If g_rst_Princi!EVALEG_APRCOM > 0 Then
      pnl_Leg_AprCom.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_APRCOM))
   End If
   
   If g_rst_Princi!EVALEG_FIRCON > 0 Then
      pnl_Leg_FirCon.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      pnl_Leg_RepLeg.Caption = Trim(g_rst_Princi!EVALEG_REPLG1)
      
      If Len(Trim(g_rst_Princi!EVALEG_REPLG2)) > 0 Then
         pnl_Leg_RepLeg.Caption = pnl_Leg_RepLeg.Caption & " / " & Trim(g_rst_Princi!EVALEG_REPLG2)
      End If
      pnl_Leg_Notari.Caption = moddat_gf_Consulta_ParDes("509", Trim(g_rst_Princi!EVALEG_BLQNOT))
   End If
   
   If g_rst_Princi!EVALEG_BLQFEC > 0 Then
      pnl_Leg_FecBlq.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_BLQFEC))
      
      If g_rst_Princi!EVALEG_TIPDOC = 1 Or g_rst_Princi!EVALEG_TIPDOC = 2 Then
         pnl_Leg_DocReg.Caption = Trim(moddat_gf_Consulta_ParDes("026", CStr(g_rst_Princi!EVALEG_TIPDOC)))
         pnl_Leg_DocReg.Caption = pnl_Leg_DocReg.Caption & " NRO.: " & Trim(g_rst_Princi!EVALEG_PARFIC) & " - ASIENTO: " & Trim(g_rst_Princi!EVALEG_NUMASI)
      Else
         pnl_Leg_DocReg.Caption = "TOMO: " & Trim(g_rst_Princi!EVALEG_BLQTOM) & " - " & "FOJAS: " & Trim(g_rst_Princi!EVALEG_BLQFOJ) & " - " & "LIBRO: " & Trim(g_rst_Princi!EVALEG_BLQLIB)
      End If
      
      txt_Leg_ObsBlq.Text = Trim(g_rst_Princi!EVALEG_BLQOBS & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatSeg()
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
     
   'Cargar Datos de Evaluación
   pnl_Seg_InfPre.Caption = Trim(g_rst_Princi!EVASEG_INFPRE & "")
   pnl_Seg_EvaPre.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAPRE))
   pnl_Seg_AplPre.Caption = moddat_gf_Consulta_ParDes("227", g_rst_Princi!EVASEG_TAPPRE)
   pnl_Seg_FoiPre.Caption = Format(g_rst_Princi!EVASEG_TASPRE, "###,###,##0.000000000") & " "
   
   pnl_Seg_InfViv.Caption = Trim(g_rst_Princi!EVASEG_INFVIV & "")
   pnl_Seg_EvaViv.Caption = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
   pnl_Seg_AplViv.Caption = moddat_gf_Consulta_ParDes("227", g_rst_Princi!EVASEG_TAPVIV)
   pnl_Seg_FoiViv.Caption = Format(g_rst_Princi!EVASEG_TASVIV, "###,###,##0.000000000") & " "
   
   txt_Seg_ObsEva.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Obteniendo Información de Póliza de Seguro
   g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
   g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   pnl_Seg_PolTit.Caption = Trim(g_rst_Princi!POLIZA_NUMDES & "")
   pnl_Seg_PolCyg.Caption = Trim(g_rst_Princi!POLIZA_NUMCYG & "")
   pnl_Seg_EmiTit.Caption = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMDES))

   pnl_Seg_PolViv.Caption = Trim(g_rst_Princi!POLIZA_NUMVIV & "")
   pnl_Seg_EmiViv.Caption = gf_FormatoFecha(CStr(g_rst_Princi!POLIZA_FEMVIV))
   
   txt_Seg_ObsPol.Text = Trim(g_rst_Princi!POLIZA_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatCof()
   g_str_Parame = "SELECT * FROM TRA_DETCOF WHERE "
   g_str_Parame = g_str_Parame & "DETCOF_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   
   g_rst_Princi.MoveFirst
   
   pnl_Cof_NumCar.Caption = Trim(g_rst_Princi!DETCOF_NUMCAR)
   pnl_Cof_NumOpe.Caption = Trim(g_rst_Princi!DETCOF_NUMOPE)
   pnl_Cof_Import.Caption = Format(g_rst_Princi!DETCOF_IMPORT, "###,###,#0.00") & " "
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   g_str_Parame = "SELECT * FROM TRA_CARCOF WHERE "
   g_str_Parame = g_str_Parame & "CARCOF_NUMCAR = '" & pnl_Cof_NumCar.Caption & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
      
   pnl_Cof_FecEmi.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CARCOF_FECEMI))
   pnl_Cof_FecVal.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CARCOF_FECVAL))
   pnl_Cof_NomBan.Caption = moddat_gf_Consulta_ParDes("505", g_rst_Princi!CARCOF_CODBAN)
   pnl_Cof_NumCta.Caption = Trim(g_rst_Princi!CARCOF_NUMCTA & "")
   
   pnl_Cof_TipMon.Caption = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CARCOF_TIPMON))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatInm()
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON))

   pnl_Inm_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMERO)
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_INTDPT))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_NOMZON))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_NOMZON) & Chr(13) & Chr(10)
   Else
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
   
   pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart

   
   pnl_Inm_TipPro.Caption = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!SOLINM_TIPPER))
   
   If g_rst_Princi!SOLINM_TIPPER = 2 Then
      'Persona Jurídica
      
      pnl_Inm_JurEmp.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PRORZS)
      pnl_Inm_JurRep.Caption = Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_PROTVI))
      r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_PROTZO))
   
      pnl_Inm_JurDir.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_PRONVI) & " " & Trim(g_rst_Princi!SOLINM_PRONUM)
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PROINT))) > 0 Then
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & " (" & Trim(g_rst_Princi!SOLINM_PROINT) & ")"
      End If
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PRONZO))) > 0 Then
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_PRONZO) & Chr(13) & Chr(10)
      Else
         pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & Chr(13) & Chr(10)
      End If
      
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 2) & "0000")
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 4) & "00")
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_PROUBI))
      
      pnl_Inm_JurDir.Caption = pnl_Inm_JurDir.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   Else
      'Persona Natural
      
      pnl_Inm_NatTit.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
      
      If g_rst_Princi!SOLINM_CYGTDO > 0 Then
         pnl_Inm_NatCyg.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_CYGTDO)) & "-" & Trim(g_rst_Princi!SOLINM_CYGNDO) & " / " & Trim(g_rst_Princi!SOLINM_CYGAPP) & " " & Trim(g_rst_Princi!SOLINM_CYGAPM) & " " & Trim(g_rst_Princi!SOLINM_CYGNOM)
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

