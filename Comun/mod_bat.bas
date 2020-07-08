Attribute VB_Name = "modbat"
Option Explicit

Public Type modbat_tpo_EstCta_Pagos
   EstCta_Pagos_FecMov  As String
   EstCta_Pagos_NumMov  As String
   EstCta_Pagos_FecPag  As String
   EstCta_Pagos_TipPag  As String
   EstCta_Pagos_NomBan  As String
   EstCta_Pagos_NumCta  As String
   EstCta_Pagos_NumCom  As String
   EstCta_Pagos_Import  As Double
End Type
Public modbat_arr_EstCta_Pagos() As modbat_tpo_EstCta_Pagos

Public Type modbat_tpo_EstCta_DetPag
   EstCta_DetPag_FecMov    As String
   EstCta_DetPag_NumMov    As String
   EstCta_DetPag_NumCuo    As String
   EstCta_DetPag_FecVct    As String
   EstCta_DetPag_Capita    As Double
   EstCta_DetPag_Intere    As Double
   EstCta_DetPag_SegPre    As Double
   EstCta_DetPag_SegViv    As Double
   EstCta_DetPag_OtrCar    As Double
   EstCta_DetPag_IntMor    As Double
   EstCta_DetPag_IntCom    As Double
   EstCta_DetPag_GasCob    As Double
   EstCta_DetPag_OtrGas    As Double
   EstCta_DetPag_TotPag    As Double
End Type
Public modbat_arr_EstCta_DetPag() As modbat_tpo_EstCta_DetPag

Public Type modbat_tpo_EstCta_ResCuo
   EstCta_ResCuo_NumCuo    As String
   EstCta_ResCuo_FecVct    As String
   EstCta_ResCuo_Capita    As Double
   EstCta_ResCuo_Intere    As Double
   EstCta_ResCuo_SegPre    As Double
   EstCta_ResCuo_SegViv    As Double
   EstCta_ResCuo_OtrCar    As Double
   EstCta_ResCuo_TotCuo    As Double
   EstCta_ResCuo_IntMor    As Double
   EstCta_ResCuo_IntCom    As Double
   EstCta_ResCuo_GasCob    As Double
   EstCta_ResCuo_OtrGas    As Double
   EstCta_ResCuo_TotPag    As Double
   EstCta_ResCuo_FecPag    As String
   EstCta_ResCuo_Situac    As String
End Type
Public modbat_arr_EstCta_ResCuo() As modbat_tpo_EstCta_ResCuo

Public Type modbat_tpo_EstCta_Cronog
   EstCta_Cronog_NumCuo    As String
   EstCta_Cronog_FecVct    As String
   EstCta_Cronog_DiaAtr    As Integer
   EstCta_Cronog_Capita    As Double
   EstCta_Cronog_Intere    As Double
   EstCta_Cronog_SegPre    As Double
   EstCta_Cronog_SegViv    As Double
   EstCta_Cronog_OtrCar    As Double
   EstCta_Cronog_TotCuo    As Double
   EstCta_Cronog_IntMor    As Double
   EstCta_Cronog_IntCom    As Double
   EstCta_Cronog_GasCob    As Double
   EstCta_Cronog_OtrGas    As Double
   EstCta_Cronog_TotDeu    As Double
   EstCta_Cronog_Situac    As String
End Type
Public modbat_arr_EstCta_Cronog() As modbat_tpo_EstCta_Cronog


