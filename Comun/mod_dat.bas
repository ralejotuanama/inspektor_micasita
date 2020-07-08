Attribute VB_Name = "moddat"
Option Explicit

Public Type moddat_tpo_Genera
   Genera_Codigo  As String
   Genera_Nombre  As String
   Genera_TipPar  As Integer
   Genera_TipVal  As Integer
   Genera_Cantid  As Double
   Genera_ValMin  As Double
   Genera_ValMax  As Double
   Genera_PlzMin  As Double
   Genera_PlzMax  As Double
   Genera_TipMon  As Integer
   Genera_Prefij  As String
   Genera_VenTDo  As Integer
   Genera_VenNDo  As String
   Genera_ConTDo  As Integer
   Genera_ConNDo  As String
   Genera_FlgAso  As Integer
   Genera_TipVia  As Integer
   Genera_NomVia  As String
   Genera_NumVia  As String
   Genera_IntDpt  As String
   Genera_TipZon  As Integer
   Genera_NomZon  As String
   Genera_Refere  As String
   Genera_UbiGeo  As String
   Genera_NumSol  As String
   Genera_TipDoc  As Integer
   Genera_NumDoc  As String
   Genera_NomCli  As String
   Genera_ConHip  As String
   Genera_EjeSeg  As String
   Genera_CodIns  As Integer
End Type

Public Type moddat_tpo_MatOpc
   MatOpc_CodMen  As String
   MatOpc_CodSub  As String
   MatOpc_Descri  As String
   MatOpc_Flag01  As String
   MatOpc_Flag02  As String
   MatOpc_Flag03  As String
   MatOpc_Flag04  As String
   MatOpc_Flag05  As String
   MatOpc_Flag06  As String
   MatOpc_Flag07  As String
   MatOpc_Flag08  As String
   MatOpc_Flag09  As String
   MatOpc_Flag10  As String
   MatOpc_Flag11  As String
   MatOpc_Flag12  As String
   MatOpc_CodT01  As String
   MatOpc_DesT01  As String
   MatOpc_CodT02  As String
   MatOpc_DesT02  As String
   MatOpc_CodT03  As String
   MatOpc_DesT03  As String
   MatOpc_CodT04  As String
   MatOpc_DesT04  As String
   MatOpc_CodT05  As String
   MatOpc_DesT05  As String
   MatOpc_CodT06  As String
   MatOpc_DesT06  As String
   MatOpc_CodT07  As String
   MatOpc_DesT07  As String
   MatOpc_CodT08  As String
   MatOpc_DesT08  As String
   MatOpc_CodT09  As String
   MatOpc_DesT09  As String
   MatOpc_CodT10  As String
   MatOpc_DesT10  As String
   MatOpc_CodT11  As String
   MatOpc_DesT11  As String
   MatOpc_CodT12  As String
   MatOpc_DesT12  As String
End Type

Public Type moddat_g_tpo_DatCli
   DatCli_TipDoc           As Integer
   DatCli_NumDoc           As String
   DatCli_DocAlt           As Integer
   DatCli_TDoAlt           As Integer
   DatCli_NDoAlt           As String
   DatCli_ApePat           As String
   DatCli_ApeMat           As String
   DatCli_ApeCas           As String
   DatCli_Nombre           As String
   DatCli_FecNac           As String
   DatCli_Paises           As String
   DatCli_UbiGeo           As String
   DatCli_NivEst           As Integer
   DatCli_Profes           As String
   DatCli_Celula           As String
   DatCli_DirEle           As String
   DatCli_ChkEle           As Integer
   DatCli_ClaSbs           As String
   DatCli_ClasMC           As String
   DatCli_ActEco           As Integer
End Type

Public Type moddat_g_tpo_ActEco
   ActEco_OrdAct           As Integer
   ActEco_TipAct           As Integer
   ActEco_Dep_TipDoc       As Integer
   ActEco_Dep_NumDoc       As String
   ActEco_Dep_FlgEmp       As Integer
   ActEco_Dep_RazSoc       As String
   ActEco_Dep_NomCom       As String
   ActEco_Dep_TipOfi       As Integer
   ActEco_Dep_SitTra       As Integer
   ActEco_Dep_TipVia       As Integer
   ActEco_Dep_NomVia       As String
   ActEco_Dep_NumVia       As String
   ActEco_Dep_IntDpt       As String
   ActEco_Dep_TipZon       As Integer
   ActEco_Dep_NomZon       As String
   ActEco_Dep_UbiGeo       As String
   ActEco_Dep_Refere       As String
   ActEco_Dep_Telef1       As String
   ActEco_Dep_Telef2       As String
   ActEco_Dep_NumFax       As String
   ActEco_Dep_CodCiu       As Integer
   ActEco_Dep_TeleRH       As String
   ActEco_Dep_AnexRH       As String
   ActEco_Dep_IngNet       As Double
   ActEco_Dep_FreHab       As Integer
   ActEco_Dep_FecIng       As String
   ActEco_Dep_CodCar       As String
   ActEco_Dep_NomCar       As String
   ActEco_Dep_NomAre       As String
   ActEco_Dep_NumAnx       As String
   ActEco_Dep_TelDir       As String
   ActEco_Dep_Celula       As String
   ActEco_Dep_DirEle       As String
   ActEco_Dep_TraAnt       As Integer
   ActEco_Dep_TipDoc_Ant   As Integer
   ActEco_Dep_NumDoc_Ant   As String
   ActEco_Dep_FlgEmp_Ant   As String
   ActEco_Dep_RazSoc_Ant   As String
   ActEco_Dep_NomCom_Ant   As String
   ActEco_Dep_Telef1_Ant   As String
   ActEco_Dep_Telef2_Ant   As String
   ActEco_Dep_FecIng_Ant   As String
   ActEco_Dep_FecCes_Ant   As String
   ActEco_Ind_TipDoc       As Integer
   ActEco_Ind_NumDoc       As String
   ActEco_Ind_TipVia       As Integer
   ActEco_Ind_NomVia       As String
   ActEco_Ind_NumVia       As String
   ActEco_Ind_IntDpt       As String
   ActEco_Ind_TipZon       As Integer
   ActEco_Ind_NomZon       As String
   ActEco_Ind_UbiGeo       As String
   ActEco_Ind_Refere       As String
   ActEco_Ind_Telef1       As String
   ActEco_Ind_Telef2       As String
   ActEco_Ind_NumFax       As String
   ActEco_Ind_CodCiu       As Integer
   ActEco_Ind_IngNet       As Double
   ActEco_Ind_IniAct       As String
   ActEco_Ind_ConLoc       As Integer
   ActEco_Ind_TipDoc_Emp   As Integer
   ActEco_Ind_NumDoc_Emp   As String
   ActEco_Ind_FlgEmp       As String
   ActEco_Ind_RazSoc_Emp   As String
   ActEco_Ind_NomCom_Emp   As String
   ActEco_Ind_Telef1_Emp   As String
   ActEco_Ind_Telef2_Emp   As String
   ActEco_Ind_FecIng_Emp   As String
   ActEco_Ind_CodCar       As String
   ActEco_Ind_NomCar       As String
   ActEco_Com_TipDoc       As Integer
   ActEco_Com_NumDoc       As String
   ActEco_Com_RazSoc       As String
   ActEco_Com_NomCom       As String
   ActEco_Com_TipVia       As Integer
   ActEco_Com_NomVia       As String
   ActEco_Com_NumVia       As String
   ActEco_Com_IntDpt       As String
   ActEco_Com_TipZon       As Integer
   ActEco_Com_NomZon       As String
   ActEco_Com_UbiGeo       As String
   ActEco_Com_Refere       As String
   ActEco_Com_Telef1       As String
   ActEco_Com_Telef2       As String
   ActEco_Com_NumFax       As String
   ActEco_Com_CodCiu       As Integer
   ActEco_Com_GirCom       As String
   ActEco_Com_IngNet       As Double
   ActEco_Com_VtaMen       As Double
   ActEco_Com_IniOpe       As String
   ActEco_Com_CodCar       As String
   ActEco_Com_NomCar       As String
   ActEco_Com_RegTri       As Integer
   ActEco_Com_PorPar       As Double
   ActEco_Com_TipLoc       As Integer
   ActEco_Com_AlqMen       As Double
   ActEco_Com_NomArr       As String
   ActEco_Com_TelArr       As String
   ActEco_Acc_TipDoc       As Integer
   ActEco_Acc_NumDoc       As String
   ActEco_Acc_FlgEmp       As Integer
   ActEco_Acc_RazSoc       As String
   ActEco_Acc_NomCom       As String
   ActEco_Acc_TipVia       As Integer
   ActEco_Acc_NomVia       As String
   ActEco_Acc_NumVia       As String
   ActEco_Acc_IntDpt       As String
   ActEco_Acc_TipZon       As Integer
   ActEco_Acc_NomZon       As String
   ActEco_Acc_UbiGeo       As String
   ActEco_Acc_Refere       As String
   ActEco_Acc_Telef1       As String
   ActEco_Acc_Telef2       As String
   ActEco_Acc_NumFax       As String
   ActEco_Acc_CodCiu       As Integer
   ActEco_Acc_IngNet       As Double
   ActEco_Acc_PorPar       As Double
   ActEco_Acc_FecAnt       As String
   ActEco_Ren_IngNet       As Double
   ActEco_Ren_Direc1       As String
   ActEco_Ren_NomAr1       As String
   ActEco_Ren_IniAl1       As String
   ActEco_Ren_Tele11       As String
   ActEco_Ren_Tele21       As String
   ActEco_Ren_AlqMe1       As Double
   ActEco_Ren_SegPro       As Integer
   ActEco_Ren_Direc2       As String
   ActEco_Ren_NomAr2       As String
   ActEco_Ren_IniAl2       As String
   ActEco_Ren_Tele12       As String
   ActEco_Ren_Tele22       As String
   ActEco_Ren_AlqMe2       As Double
   ActEco_Otr_IngNet       As Double
   ActEco_Otr_Activi       As String
   ActEco_Otr_CodCiu       As Integer
   ActEco_Otr_Observ       As String
End Type

Public Type moddat_g_tpo_DatCyg
   DatCyg_TipDoc        As Integer
   DatCyg_NumDoc        As String
   DatCyg_ApePat        As String
   DatCyg_ApeMat        As String
   DatCyg_ApeCas        As String
   DatCyg_Nombre        As String
   DatCyg_FecNac        As String
   DatCyg_Paises        As String
   DatCyg_DptNac        As String
   DatCyg_PrvNac        As String
   DatCyg_DstNac        As String
   DatCyg_NivEst        As Integer
   DatCyg_Profes        As String
   DatCyg_Celula        As String
   DatCyg_DirEle        As String
   DatCyg_AutEnv        As Integer
   DatCyg_ClaSbs        As String
   DatCyg_ClasMc        As String
   DatCyg_RegAct        As Integer
End Type

Public Type moddat_g_tpo_DatCom
   DatCom_TipMon       As Integer
   DatCom_CodMod       As String
   DatCom_TasInt       As Double
   DatCom_ComVta_Dol   As Double
   DatCom_MtoInm_Dol   As Double
   DatCom_MtoEst_Dol   As Double
   DatCom_ApoPro_Dol   As Double
   DatCom_ComVta_Sol   As Double
   DatCom_MtoInm_Sol   As Double
   DatCom_MtoEst_Sol   As Double
   DatCom_ApoPro_Sol   As Double
   DatCom_FmvBbp_Sol   As Double
   DatCom_MefPbp_Sol   As Double
   DatCom_MtoAFP_Sol   As Double
   DatCom_MtoBMS_Sol   As Double
   DatCom_PreMto_Sol   As Double
   DatCom_MtoGCi_Sol   As Double
   DatCom_MtoPre_Sol   As Double
   DatCom_MtoPre_Dol   As Double
   DatCom_PerGra       As Integer
   DatCom_PlaAno       As Integer
   DatCom_CuoExt       As Integer
   DatCom_DiaPag       As Integer
   DatCom_IntGra       As Double
   DatCom_Observ       As String
   DatCom_CodPrd       As String
   DatCom_CodSub       As String
   DatCom_EjeSeg       As String
   DatCom_ConHip       As String
   DatCom_FecSol       As String
   DatCom_CodIns       As Integer
   DatCom_TipEva       As Integer
   DatCom_MtoPre_Mpr   As Double
   DatCom_PriViv       As Integer
   DatCom_EsgDes       As String
   DatCom_ComVta_Mon   As Integer
   DatCom_MtoPre_Cal   As Double
   DatCom_PlaAno_Cal   As Integer
   DatCom_PerGra_Cal   As Integer
   DatCom_CuoExt_Cal   As Integer
   DatCom_TipSeg_Cal   As Integer
   DatCom_MesAho       As Integer
   DatCom_TasEsp       As Integer
   DatCom_BmsTas       As Double
End Type

Public moddat_g_arr_ActEco_Tit(2)   As moddat_g_tpo_ActEco
Public moddat_g_arr_ActEco_Cyg(2)   As moddat_g_tpo_ActEco
Public moddat_g_arr_DatCyg(1)       As moddat_g_tpo_DatCli
Public moddat_g_arr_Genera()        As moddat_tpo_Genera
Public moddat_g_arr_GenAux()        As moddat_tpo_Genera

Public moddat_g_int_FlgGrb          As Integer
Public moddat_g_int_FlgAct          As Integer
Public moddat_g_int_FlgGrb_1        As Integer
Public moddat_g_int_FlgAct_1        As Integer
Public moddat_g_int_FlgGrb_2        As Integer
Public moddat_g_int_FlgAct_2        As Integer

Public moddat_g_int_TipDoc          As Integer
Public moddat_g_str_TipDoc          As String
Public moddat_g_str_NumDoc          As String
Public moddat_g_str_NomCli          As String
Public moddat_g_int_CygTDo          As Integer
Public moddat_g_str_CygTDo          As String
Public moddat_g_str_CygNDo          As String
Public moddat_g_str_CygNom          As String
Public moddat_g_int_TipCli          As Integer
Public moddat_g_int_OrdAct          As Integer

Public moddat_g_int_CygTDo_1        As Integer
Public moddat_g_str_CygNDo_1        As String
Public moddat_g_str_CygNom_1        As String
Public moddat_g_int_FlgCyg          As Integer
Public moddat_g_int_FlgGOK          As Integer
Public moddat_g_int_CntErr          As Integer

Public moddat_g_str_Codigo          As String
Public moddat_g_str_Descri          As String
Public moddat_g_str_CodIte          As String
Public moddat_g_str_DesIte          As String
Public moddat_g_str_TipPar          As String
Public moddat_g_str_CodMod          As String
Public moddat_g_str_DesMod          As String
Public moddat_g_str_CodGrp          As String
Public moddat_g_str_DesGrp          As String

Public moddat_g_int_EdaAno          As Integer
Public moddat_g_int_EdaMes          As Integer
Public moddat_g_str_CodPrd          As String
Public moddat_g_str_NomPrd          As String
Public moddat_g_str_CodSub          As String
Public moddat_g_str_DesSub          As String
Public moddat_g_str_ClaPrd          As String
Public moddat_g_str_CodCam          As String
Public moddat_g_str_TipCre          As String
Public moddat_g_str_SitCre          As String
Public moddat_g_str_ClaGar          As String

Public moddat_g_str_NumSol          As String
Public moddat_g_str_NumOpe          As String
Public moddat_g_str_CodEje          As String
Public moddat_g_str_EjeVta          As String
Public moddat_g_str_CodEjeSeg       As String
Public moddat_g_str_NomEjeSeg       As String
Public moddat_g_str_CodConHip       As String
Public moddat_g_str_NomConHip       As String

Public moddat_g_str_Moneda          As String
Public moddat_g_str_FecIng          As String
Public moddat_g_str_FecApr          As String
Public moddat_g_str_FecDes          As String
Public moddat_g_str_FecRec          As String
Public moddat_g_str_FecAnu          As String
Public moddat_g_int_SitIns          As Integer
Public moddat_g_int_Situac          As Integer
Public moddat_g_int_InmIde          As Integer
Public moddat_g_int_InsAct          As Integer
Public moddat_g_int_MotRec          As Integer
Public moddat_g_int_TipRep          As Integer
Public moddat_g_int_TipObs          As Integer
Public moddat_g_str_Observ          As String
Public moddat_g_int_TipAut          As Integer
Public moddat_g_int_MotExc          As Integer  'LMD13102011
Public moddat_g_int_TipMon          As Integer
Public moddat_g_int_CodIns          As Integer
Public moddat_g_str_InsAct          As String
Public moddat_g_str_Situac          As String
Public moddat_g_str_FecSis          As String
Public moddat_g_str_HorSis          As String
Public moddat_g_int_NumCuo          As Integer
Public moddat_g_int_MonDes          As Integer
Public moddat_g_int_MonViv          As Integer
Public moddat_g_int_MonOtr          As Integer
Public moddat_g_dbl_MtoPre          As Double
Public moddat_g_dbl_SalCap          As Double
Public moddat_g_int_MtoAfp          As Double
Public moddat_g_int_TotCuo          As Integer
Public moddat_g_int_CuoPen          As Integer
Public moddat_g_int_TDoEmp          As Integer
Public moddat_g_str_NDoEmp          As String
Public moddat_g_str_CodEmp          As String
Public moddat_g_str_RazSoc          As String
Public moddat_g_str_NomCom          As String
Public moddat_g_str_Direcc          As String
Public moddat_g_str_Distri          As String
Public moddat_g_str_EmpSegDes       As String
Public moddat_g_str_EmpSegViv       As String
Public moddat_g_int_TipPan          As Integer
Public moddat_g_dbl_IngDec          As Double
Public moddat_g_int_TipRec          As Integer
Public moddat_g_str_FecIni          As String
Public moddat_g_str_FecFin          As String
Public moddat_g_str_FecCan          As String
Public moddat_g_int_TipGar          As Integer
Public moddat_g_str_FecHip          As String
Public moddat_g_str_BanFia          As String
Public moddat_g_str_NumFia          As String
Public moddat_g_str_FecFia          As String
Public moddat_g_int_TipEva          As Integer
Public moddat_g_dbl_TasInt          As Double
Public moddat_g_str_CodGen          As String
Public moddat_g_str_DesGen          As String
Public moddat_g_str_CodMes          As Integer
Public moddat_g_str_CodAno          As Integer
Public moddat_g_int_NumObs          As Integer
Public moddat_g_str_DesObs          As String
Public moddat_g_dbl_TotGar          As Double

Public moddat_g_int_EstCiv          As Integer
Public moddat_g_int_RegCyg          As Integer
Public moddat_g_int_ComRta          As Integer
Public moddat_g_str_FecNac_Tit      As String
Public moddat_g_str_FecNac_Cyg      As String
Public moddat_g_str_UbiGeo          As String
Public moddat_g_int_FlgCre          As Integer
Public moddat_g_int_FlgPre          As Integer
Public moddat_g_int_FlgActPri_Cli   As Integer
Public moddat_g_int_FlgActSec_Cli   As Integer
Public moddat_g_int_FlgActPri_Cyg   As Integer
Public moddat_g_int_FlgActSec_Cyg   As Integer
Public moddat_g_int_FlgActEnv       As Integer

Public moddat_g_wsp_Access          As Workspace
Public moddat_g_bdt_Report          As Database
Public moddat_g_rst_Access          As DAO.Recordset
Public moddat_g_rst_RecDAO          As DAO.Recordset
Public moddat_g_str_CadDAO          As String

Public moddat_g_str_AgrCRC          As String
Public moddat_g_str_AgrCME          As String
Public moddat_g_str_AgrCOF          As String
Public moddat_g_str_AgrTMIC         As String
Public moddat_g_str_Agr1MIC         As String
Public moddat_g_str_Agr2MIC         As String
Public moddat_g_str_AgrTFMV         As String
Public moddat_g_str_AgrMIHG         As String
Public moddat_g_str_Agr1FMV         As String
Public moddat_g_str_Agr2FMV         As String

Public moddat_g_dbl_ValSGa          As Double
Public moddat_g_dbl_ValGHi          As Double
Public moddat_g_dbl_ValGLi          As Double

Public moddat_g_dbl_ValNv1          As Double
Public moddat_g_dbl_ValNv2          As Double
Public moddat_g_dbl_ValNv3          As Double


        
        
        
        




Public Sub moddat_gs_Carga_Cbr_ResAcc(p_Combo As ComboBox, ByVal p_TipAcc As Integer)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '303' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 AND "
   Select Case p_TipAcc
      Case 1:  g_str_Parame = g_str_Parame & "PARDES_CODITE >= 1001 AND PARDES_CODITE <= 1999 "
      Case 2:  g_str_Parame = g_str_Parame & "PARDES_CODITE >= 2001 AND PARDES_CODITE <= 2999 "
   End Select
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub



Public Function validacionclientexproblemas(ByVal vara1 As String, ByVal vara2 As String) As String
    Dim http As New ChilkatHttp
    Dim success1 As Long
    Dim MyPos As Integer
      '  Any string unlocks the component for the 1st 30-days.
    success1 = http.UnlockComponent("Anything for 30-day trial")
    If (success1 <> 1) Then
    Debug.Print "Error"
    End If
     'MsgBox (success1)
    Dim var1, var2 As String
    var1 = vara1
    var2 = vara2

    Dim soapXml As New ChilkatXml
    soapXml.Tag = "soap12:Envelope"
    Dim success As Long
    success = soapXml.AddAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    success = soapXml.AddAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    success = soapXml.AddAttribute("xmlns:soap12", "http://www.w3.org/2003/05/soap-envelope")
    soapXml.NewChild2 "soap12:Body", ""
    success = soapXml.GetChild2(0)
    soapXml.NewChild2 "LoadWSInspektor", ""
    success = soapXml.GetChild2(0)
    success = soapXml.AddAttribute("xmlns", "http://tempuri.org/")
    soapXml.NewChild2 "Numeiden", var1
    soapXml.NewChild2 "Nombre", var2
    soapXml.NewChild2 "Password", "M1c4s1t4.2020"
    soapXml.GetRoot2

    Debug.Print soapXml.GetXml()

    Dim req As New ChilkatHttpRequest
    req.HttpVerb = "POST"
    req.SendCharset = 0
    req.AddHeader "Content-Type", "application/soap+xml; charset=utf-8"
    req.AddHeader "SOAPAction", "http://tempuri.org/LoadWSInspektor"
    req.Path = "https://inspektortest.datalaft.com:8749/WSInspektor.asmx"
    success = req.LoadBodyFromString(soapXml.GetXml(), "utf-8")
    http.FollowRedirects = 1
    Dim resp As ChilkatHttpResponse
    Set resp = http.SynchronousRequest("wsf.cdyne.com", 80, 0, req)
    If (http.LastMethodSuccess = 0) Then
      Debug.Print http.LastErrorText
    Else
      Dim xmlResponse As New ChilkatXml
      success = xmlResponse.LoadXml(resp.BodyStr)
      Debug.Print xmlResponse.GetXml()

    End If
    Dim responseStatusCode As Long
     http.SetRequestHeader "SOAPAction", "http://tempuri.org/LoadWSInspektor"
     http.SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
     Dim endPoint As String
     endPoint = "https://inspektortest.datalaft.com:8749/WSInspektor.asmx"
     Set resp = http.PostXml(endPoint, soapXml.GetXml(), "utf-8")
    If (resp Is Nothing) Then
    Else
    responseStatusCode = resp.StatusCode
    'You may wish to verify that the responseStatusCode equals 200...

    'You may examine the exact HTTP header sent with the POST like this:

 
    'Examine the XML returned by the web service:
    'Debug.Print (Text1.Text & "XML Response:" & vbCrLf)
     Dim xmlResp As New ChilkatXml
    
    success = xmlResp.LoadXml(resp.BodyStr)
    
    'Assume the LoadXml is successful...
    'Get rid of the SOAP wrappings and get to the meat of this particular response.
    'The TagContent method returns the content of the 1st node in the XML document
    'having a specific tag:
     Dim unwrappedXml As String
     unwrappedXml = xmlResp.TagContent("LoadWSInspektorResult")
     'Text1.Text = Text1.Text & unwrappedXml & vbCrLf

     'The unwrapped XML could be loaded into an XML object and parsed...
     'Dim xmlMeat As New ChilkatXml
   
     modgen_g_str_rptwebservice = unwrappedXml
     MyPos = InStr(unwrappedXml, "No existen registros asociados")
    End If
    validacionclientexproblemas = MyPos
End Function
  


  
    
   

 
 
 
 
 





Public Sub moddat_gs_Carga_Combo_ParPrd(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodGrp As String)
   ReDim p_Arregl(0)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_FecSis()
   moddat_g_str_FecSis = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Fecha del Sistema
   g_str_Parame = "SELECT TO_CHAR(SYSDATE,'dd/mm/yyyy') AS VS_FECSIS FROM DUAL"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   moddat_g_str_FecSis = g_rst_Genera!VS_FECSIS
   DoEvents
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   'Obteniendo Hora del Sistema
   g_str_Parame = "SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') AS VS_HORSIS FROM DUAL"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   moddat_g_str_HorSis = g_rst_Genera!VS_HORSIS
   DoEvents
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Proyec(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = "SELECT * FROM PRY_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_TITULO"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!DATGEN_TITULO)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!DATGEN_TITULO)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!DATGEN_PRYMCS
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = 0
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_PryVin(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = "SELECT * FROM PRY_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_PRYMCS = 1 AND "
   g_str_Parame = g_str_Parame & "DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_TITULO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      If g_rst_Genera!DATGEN_VENTDO <> g_rst_Genera!DATGEN_CONTDO And g_rst_Genera!DATGEN_VENNDO <> g_rst_Genera!DATGEN_CONNDO Then
         p_Combo.AddItem Trim(g_rst_Genera!DATGEN_TITULO) & " (" & moddat_gf_Consulta_RazSoc(CStr(g_rst_Genera!DATGEN_VENTDO), Trim(g_rst_Genera!DATGEN_VENNDO)) & " - " & moddat_gf_Consulta_RazSoc(CStr(g_rst_Genera!DATGEN_CONTDO), Trim(g_rst_Genera!DATGEN_CONNDO)) & ")"
      Else
         p_Combo.AddItem Trim(g_rst_Genera!DATGEN_TITULO) & " (" & moddat_gf_Consulta_RazSoc(CStr(g_rst_Genera!DATGEN_VENTDO), Trim(g_rst_Genera!DATGEN_VENNDO)) & ")"
      End If
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!DATGEN_TITULO)
      p_Arregl(UBound(p_Arregl)).Genera_VenTDo = g_rst_Genera!DATGEN_VENTDO
      p_Arregl(UBound(p_Arregl)).Genera_VenNDo = Trim(g_rst_Genera!DATGEN_VENNDO)
      p_Arregl(UBound(p_Arregl)).Genera_ConTDo = g_rst_Genera!DATGEN_CONTDO
      p_Arregl(UBound(p_Arregl)).Genera_ConNDo = Trim(g_rst_Genera!DATGEN_CONNDO)
      p_Arregl(UBound(p_Arregl)).Genera_FlgAso = g_rst_Genera!DATGEN_FLGCON
      p_Arregl(UBound(p_Arregl)).Genera_TipVia = g_rst_Genera!DatGen_TipVia
      p_Arregl(UBound(p_Arregl)).Genera_NomVia = Trim(g_rst_Genera!DatGen_NomVia)
      p_Arregl(UBound(p_Arregl)).Genera_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
      p_Arregl(UBound(p_Arregl)).Genera_IntDpt = Trim(g_rst_Genera!DATGEN_INTDPT & "")
      p_Arregl(UBound(p_Arregl)).Genera_TipZon = g_rst_Genera!DatGen_TipZon
      p_Arregl(UBound(p_Arregl)).Genera_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
      p_Arregl(UBound(p_Arregl)).Genera_Refere = Trim(g_rst_Genera!DATGEN_REFERE & "")
      p_Arregl(UBound(p_Arregl)).Genera_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_PryNVi(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodBco As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM PRY_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_CODBCO = '" & p_CodBco & "' "
   g_str_Parame = g_str_Parame & "   AND DATGEN_PRYMCS = 2 "
   g_str_Parame = g_str_Parame & "   AND DATGEN_SITUAC = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY DATGEN_TITULO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!DATGEN_TITULO)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!DATGEN_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!DATGEN_TITULO)
      p_Arregl(UBound(p_Arregl)).Genera_VenTDo = g_rst_Genera!DATGEN_VENTDO
      p_Arregl(UBound(p_Arregl)).Genera_VenNDo = Trim(g_rst_Genera!DATGEN_VENNDO)
      p_Arregl(UBound(p_Arregl)).Genera_ConTDo = g_rst_Genera!DATGEN_CONTDO
      p_Arregl(UBound(p_Arregl)).Genera_ConNDo = Trim(g_rst_Genera!DATGEN_CONNDO)
      p_Arregl(UBound(p_Arregl)).Genera_FlgAso = g_rst_Genera!DATGEN_FLGCON
      p_Arregl(UBound(p_Arregl)).Genera_TipVia = g_rst_Genera!DatGen_TipVia
      p_Arregl(UBound(p_Arregl)).Genera_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
      p_Arregl(UBound(p_Arregl)).Genera_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
      p_Arregl(UBound(p_Arregl)).Genera_IntDpt = Trim(g_rst_Genera!DATGEN_INTDPT & "")
      p_Arregl(UBound(p_Arregl)).Genera_TipZon = g_rst_Genera!DatGen_TipZon
      p_Arregl(UBound(p_Arregl)).Genera_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
      p_Arregl(UBound(p_Arregl)).Genera_Refere = Trim(g_rst_Genera!DATGEN_REFERE & "")
      p_Arregl(UBound(p_Arregl)).Genera_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_NomPry(ByVal p_CodPry As String) As String
   moddat_gf_Consulta_NomPry = ""
   
   g_str_Parame = "SELECT * FROM PRY_DATGEN WHERE DATGEN_CODIGO = '" & p_CodPry & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_NomPry = Trim(g_rst_Listas!DATGEN_TITULO)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_EjecMC(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipEje As Integer, Optional ByVal p_Situac As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_EJETIP A, CRE_EJECMC B WHERE "
   g_str_Parame = g_str_Parame & "A.EJETIP_CODEJE = B.EJECMC_CODEJE AND "
   g_str_Parame = g_str_Parame & "A.EJETIP_TIPEJE = " & CStr(p_TipEje) & " AND "
   
   If p_Situac <> 2 Then
      g_str_Parame = g_str_Parame & "A.EJETIP_TIPEJE = " & CStr(p_TipEje) & " AND "
      g_str_Parame = g_str_Parame & "B.EJECMC_SITUAC = 1 "
   Else
      g_str_Parame = g_str_Parame & "A.EJETIP_TIPEJE = " & CStr(p_TipEje) & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY B.EJECMC_APEPAT ASC, B.EJECMC_APEMAT ASC, B.EJECMC_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!EJECMC_APEPAT) & " " & Trim(g_rst_Listas!EJECMC_APEMAT) & " " & Trim(g_rst_Listas!EJECMC_NOMBRE)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!EJECMC_CODEJE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = ""
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = 0
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = 0
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_TipMon(p_Combo As ComboBox, ByVal p_FlgOpc As Integer)
   'Tipo de Moneda
   'Rubro 002
   'Cuando p_FlgOpc sea
   '  1  :  Nuevos Soles / Dólares Americanos
   '  2  :  Dólares Americanos / VAC
   '  3  :  Todas las Monedas
   
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '204' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      If CInt(g_rst_Genera!PARDES_CODITE) = 1 Then
         If p_FlgOpc = 1 Or p_FlgOpc = 3 Then
            p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
         End If
      ElseIf CInt(g_rst_Genera!PARDES_CODITE) = 2 Then
         p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
         p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
      Else
         If p_FlgOpc = 2 Or p_FlgOpc = 3 Then
            p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
         End If
      End If
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   p_Combo.ListIndex = -1
End Sub

Public Sub moddat_gs_Carga_TipDocIde(p_Combo As ComboBox, ByVal p_FlgOpc As Integer)
   'Tipo de Documentos de Identidad
   'Rubro 001
   'Cuando p_FlgOpc sea
   '  1  :  Tipos de Documentos de Personas Naturales
   '  2  :  Tipos de Documentos de Personas Jurídicas
   '  3  :  Todos los Tipos de Documentos
   
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '203' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      Select Case CInt(g_rst_Genera!PARDES_CODITE)
         Case 1
            If p_FlgOpc = 1 Or p_FlgOpc = 3 Then
               p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
               p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
            End If
      
         Case 2
            If p_FlgOpc = 1 Or p_FlgOpc = 3 Then
               p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
               p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
            End If
            
         Case 3
            If p_FlgOpc = 1 Or p_FlgOpc = 3 Then
               p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
               p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
            End If
      
         Case 4
            If p_FlgOpc = 1 Or p_FlgOpc = 3 Then
               p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
               p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
            End If
            
         Case 7
            If p_FlgOpc = 2 Or p_FlgOpc = 3 Then
               p_Combo.AddItem CStr(CInt(g_rst_Genera!PARDES_CODITE)) & ": " & Trim$(g_rst_Genera!PARDES_DESCRI)
               p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
            End If
            
      End Select
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   p_Combo.ListIndex = -1
End Sub

Public Function moddat_gf_Inserta_SolDoc(ByVal p_NumSol As String, p_TipDoc As Integer, p_CodPrd As String, p_CodSub As String, p_CodAct As Integer, p_CodGrp As String, p_CodIte As String, p_FecRec As String) As Integer
   Dim r_int_Contad     As Integer

   moddat_gf_Inserta_SolDoc = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_SOLDOC_INSERTA ("
      
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & p_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodSub & "', "
      g_str_Parame = g_str_Parame & CStr(p_CodAct) & ", "
      g_str_Parame = g_str_Parame & "'" & p_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodIte & "', "
      g_str_Parame = g_str_Parame & p_FecRec & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_CRESOLDOC. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_SolDoc = True
End Function

Public Function moddat_gf_Buscar_NomCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, Optional ByVal p_FlgAll As Integer) As String
   moddat_gf_Buscar_NomCli = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      If p_FlgAll = 1 Then
         If Len(Trim(g_rst_Listas!DatGen_ApeCas)) > 0 Then
            moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DATGEN_APEPAT) & " " & Trim(g_rst_Listas!DATGEN_APEMAT) & " DE " & Trim(g_rst_Listas!DatGen_ApeCas) & " " & Trim(g_rst_Listas!DATGEN_NOMBRE)
         Else
            moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DATGEN_APEPAT) & " " & Trim(g_rst_Listas!DATGEN_APEMAT) & " " & Trim(g_rst_Listas!DATGEN_NOMBRE)
         End If
      Else
         moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DATGEN_APEPAT) & " " & Trim(g_rst_Listas!DATGEN_APEMAT) & " " & Trim(g_rst_Listas!DATGEN_NOMBRE)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Buscar_NomCli_PlanAhorro(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
   moddat_gf_Buscar_NomCli_PlanAhorro = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_AHOCLI "
   g_str_Parame = g_str_Parame & " WHERE AHOCLI_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND AHOCLI_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Buscar_NomCli_PlanAhorro = Trim(g_rst_Listas!AHOCLI_APEPAT) & " " & Trim(g_rst_Listas!AHOCLI_APEMAT) & " " & Trim(g_rst_Listas!AHOCLI_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Buscar_NomEje(ByVal p_CodEje As String) As String
   moddat_gf_Buscar_NomEje = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODEJE = '" & p_CodEje & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Buscar_NomEje = Trim(g_rst_Genera!EJECMC_APEPAT) & " " & Trim(g_rst_Genera!EJECMC_APEMAT) & " " & Trim(g_rst_Genera!EJECMC_NOMBRE)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Inserta_SegDet(ByVal p_NumSol As String, ByVal p_CodIns As Integer, ByVal p_CodOcu As Integer, ByVal p_NumObs As Integer, ByVal p_DesObs As String, ByVal p_SitObs As Integer, ByVal p_MotRec As Integer) As Integer
   moddat_gf_Inserta_SegDet = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_TRA_SEGDET ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
      g_str_Parame = g_str_Parame & CStr(p_CodOcu) & ", "                                          'Ocurrencia
      g_str_Parame = g_str_Parame & CStr(p_NumObs) & ", "                                          'Número de Observación
      g_str_Parame = g_str_Parame & "'" & p_DesObs & "', "                                         'Descripción de Observación
      g_str_Parame = g_str_Parame & CStr(p_SitObs) & ", "                                          'Situación Observación
      g_str_Parame = g_str_Parame & CStr(p_MotRec) & ", "                                          'Motivo de Rechazo
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_TRA_SEGDET. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_SegDet = True
End Function

Public Function moddat_gf_Inserta_SegExc(ByVal p_NumSol As String, p_CodIns As Integer, p_NumExc As Integer, p_Descri As String, p_TipAut As Integer) As Integer
   moddat_gf_Inserta_SegExc = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If p_CodIns <> 21 Then
      moddat_g_int_MotExc = 0
   End If
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_SEGEXC_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
      g_str_Parame = g_str_Parame & CStr(p_NumExc) & ", "                                          'Ocurrencia
      g_str_Parame = g_str_Parame & "'" & p_Descri & "', "                                         'Descripción de Observación
      g_str_Parame = g_str_Parame & CStr(p_TipAut) & ", "                                          'Situación Observación
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "'" & moddat_g_int_MotExc & "') "                              'Código Motivo de Excepción
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_SEGEXC_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_SegExc = True
End Function

Public Function moddat_gf_Inserta_AprCon(ByVal p_NumSol As String, p_CodIns As Integer, p_Descri As String) As Integer
   moddat_gf_Inserta_AprCon = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_SEGCON_INSERTA ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
      g_str_Parame = g_str_Parame & "'" & p_Descri & "', "                                         'Descripción de Observación
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_SEGCON_INSERTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_AprCon = True
End Function

Public Function moddat_gf_Inserta_LevCon(ByVal p_NumSol As String, p_CodIns As Integer, p_Descri As String) As Integer
   moddat_gf_Inserta_LevCon = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_SEGCON_LEVANTA ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
      g_str_Parame = g_str_Parame & "'" & p_Descri & "', "                                         'Descripción de Observación
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_SEGCON_LEVANTA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_LevCon = True
End Function

Public Function moddat_gf_Inserta_Seguim(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As Integer
   moddat_gf_Inserta_Seguim = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_INSERTA_TRA_SEGUIM ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                                         'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                                          'Instancia
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_INSERTA_TRA_SEGUIM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Inserta_Seguim = True
End Function

Public Function moddat_gf_Modifica_Seguim(ByVal p_NumSol As String, ByVal p_CodIns As Integer, ByVal p_DiaTra As Integer, ByVal p_Situac As Integer, ByVal p_FlgSit As Integer) As Integer
   moddat_gf_Modifica_Seguim = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MODIFICA_TRA_SEGUIM ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                   'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "                    'Código de Instancia
      g_str_Parame = g_str_Parame & CStr(p_DiaTra) & ", "                    'Días Transcurridos
      g_str_Parame = g_str_Parame & CStr(p_Situac) & ", "                     'Situación
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & CStr(p_FlgSit) & ", "                                       'Flag de Si Actualiza solo Situación
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_TRA_SEGUIM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Modifica_Seguim = True
End Function

Public Function moddat_gf_Modifica_SegDet_Observ(ByVal p_NumSol As String, ByVal p_CodIns As Integer, ByVal p_CodOcu As Integer, ByVal p_NumObs As String, ByVal p_Descar As String, ByVal p_SitObs As Integer) As Integer
   moddat_gf_Modifica_SegDet_Observ = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MODIFICA_TRA_SEGDET_OBSERV ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "                              'Número de Solicitud
      g_str_Parame = g_str_Parame & CStr(p_CodIns) & ", "
      g_str_Parame = g_str_Parame & CStr(p_CodOcu) & ", "
      g_str_Parame = g_str_Parame & CStr(CInt(p_NumObs)) & ", "
      g_str_Parame = g_str_Parame & "'" & p_Descar & "',"                                   'Descripción de Observación
      g_str_Parame = g_str_Parame & CStr(p_SitObs) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_MODIFICA_TRA_SEGDET_OBSERV. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_gf_Modifica_SegDet_Observ = True
End Function

Public Sub moddat_gs_Carga_Produc(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipCre As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE PRODUC_SITUAC = 1 "
   
   If p_TipCre <> 99 Then
      g_str_Parame = g_str_Parame & "AND PRODUC_CODCLA = " & CStr(p_TipCre) & " "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY PRODUC_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!PRODUC_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Listas!Produc_Codigo)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Listas!PRODUC_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Listas!PRODUC_CODCLA
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_Produc(ByVal p_CodPrd As String) As String
   moddat_gf_Consulta_Produc = ""
   p_CodPrd = Format(p_CodPrd, "000")

   g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE "
   g_str_Parame = g_str_Parame & "PRODUC_CODIGO = '" & p_CodPrd & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   moddat_gf_Consulta_Produc = Trim$(g_rst_Listas!PRODUC_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Obtiene_TipCam(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer) As Double
   Dim r_str_FecDia     As String
   
   moddat_gf_Obtiene_TipCam = 0
   
   Call moddat_gs_FecSis
   r_str_FecDia = Format(CDate(moddat_g_str_FecSis), "yyyymmdd")

   'Obteniendo Tipo de Cambio del Dia
   g_str_Parame = "SELECT * FROM OPE_TIPCAM WHERE "
   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = " & CStr(p_TipCam) & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA = " & r_str_FecDia & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_TIPMON = " & CStr(p_TipMon) & " ORDER BY "
   g_str_Parame = g_str_Parame & "TIPCAM_HORDIA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Obtiene_TipCam = g_rst_Genera!TIPCAM_VENTAS
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Carga_CtaBan(ByVal p_CodBan As String, p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN WHERE "
   g_str_Parame = g_str_Parame & "CTABAN_CODBAN = '" & p_CodBan & "' AND "
   g_str_Parame = g_str_Parame & "CTABAN_SITUAC = 1 ORDER BY CTABAN_NUMCTA ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!CtaBan_NumCta) & " - (" & moddat_gf_Consulta_ParDes("229", g_rst_Genera!ctaban_TipMon) & ")"
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!CtaBan_NumCta)
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!ctaban_TipMon
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_GirCom(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_GIRCOM "
   g_str_Parame = g_str_Parame & "ORDER BY GIRCOM_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!GIRCOM_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!GIRCOM_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!GIRCOM_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Listas!GIRCOM_CODCIU
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Listas!GIRCOM_MARUTI
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_EmpSeg(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, Optional p_Situac As Integer = 0)
   p_Combo.Clear
   ReDim p_Arregl(0)
   g_str_Parame = ""
   
   If p_Situac = 1 Then
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_SEGEMP "
      g_str_Parame = g_str_Parame & " WHERE SEGEMP_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "ORDER BY SEGEMP_RAZSOC ASC"
   Else
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_SEGEMP "
      g_str_Parame = g_str_Parame & "ORDER BY SEGEMP_RAZSOC ASC"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!SEGEMP_RAZSOC)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!SEGEMP_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!SEGEMP_RAZSOC)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_TipSeg(p_Combo As ComboBox, p_CodEmp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_SEGTIP "
   g_str_Parame = g_str_Parame & " WHERE SEGTIP_CODIGO = '" & p_CodEmp & "' "
   g_str_Parame = g_str_Parame & " ORDER BY SEGTIP_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim(g_rst_Genera!SEGTIP_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = g_rst_Genera!SEGTIP_TIPSEG
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_LibCon(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM CNTBL_LIBRO "
   g_str_Parame = g_str_Parame & "ORDER BY NRO_LIBRO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!NRO_LIBRO) & " - " & UCase(Trim$(g_rst_Genera!DESC_LIBRO))
      p_Combo.ItemData(p_Combo.NewIndex) = g_rst_Genera!NRO_LIBRO
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Operac(p_Arregl() As moddat_tpo_Genera, p_Combo As ComboBox)
   ReDim p_Arregl(0)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM OPERACION_TIPO "
'   g_str_Parame = g_str_Parame & "SUBSTR(OPERACION,1,1) = '0' "
   g_str_Parame = g_str_Parame & "ORDER BY OPERACION ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!OPERACION) & " - " & UCase(Trim$(g_rst_Genera!Descripcion))
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Genera!OPERACION)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Genera!Descripcion)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Produc_1(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE PRODUC_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PRODUC_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   'Agregando Registro General
   p_Combo.AddItem "<< GENERALES >>>"
   
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = "999"
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = "<< GENERALES >>>"
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PRODUC_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!Produc_Codigo)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PRODUC_DESCRI)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_MatCon(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = "SELECT * FROM MATRIZ_CNTBL_OPER "
   g_str_Parame = g_str_Parame & "ORDER BY MATRIZ_CTBL ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   'Agregando Registro General
   p_Combo.AddItem "<< NINGUNO >>>"
   
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = "000000"
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = "<< NINGUNO >>>"
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!MATRIZ_CTBL) & " - " & UCase(Trim$(g_rst_Genera!Descripcion))
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!MATRIZ_CTBL)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!Descripcion)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_Operac(ByVal p_CodPrd As String, ByVal p_CodAcc As String) As String
   moddat_gf_Consulta_Operac = ""
   p_CodPrd = Format(p_CodPrd, "000")
   p_CodAcc = Format(p_CodAcc, "000000")

   g_str_Parame = "SELECT * FROM CRE_OPERAC WHERE "
   g_str_Parame = g_str_Parame & "OPERAC_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "OPERAC_CODACC = '" & p_CodAcc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   moddat_gf_Consulta_Operac = Trim$(g_rst_Genera!OPERAC_CODIGO)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gf_RazSoc_NomCom()
   moddat_g_str_RazSoc = ""
   moddat_g_str_NomCom = ""
   
   'Buscando Razón Social y Nombre Comercial
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_str_NDoEmp & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If

   moddat_g_str_RazSoc = Trim(g_rst_Listas!DATGEN_RAZSOC)
   moddat_g_str_NomCom = Trim(g_rst_Listas!DATGEN_NOMCOM & "")
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_RazSoc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, Optional ByRef p_Telefo As String) As String
   moddat_gf_Consulta_RazSoc = ""
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If

   moddat_gf_Consulta_RazSoc = Trim(g_rst_Listas!DATGEN_RAZSOC)
   p_Telefo = Trim(g_rst_Listas!DATGEN_TELEF1 & "")
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_NomCom(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As String
   moddat_gf_Consulta_NomCom = ""
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If

   moddat_gf_Consulta_NomCom = Trim(g_rst_Listas!DATGEN_NOMCOM)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function


Public Function moddat_gf_CtaCtb(ByVal p_CodCta As String) As Integer
   moddat_gf_CtaCtb = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CNTBL_CNTA WHERE "
   g_str_Parame = g_str_Parame & "CNTA_CTBL = '" & p_CodCta & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   moddat_gf_CtaCtb = True
End Function

Public Function moddat_gf_Consulta_EmpSeg(ByVal p_TipSeg As Integer) As String
   moddat_gf_Consulta_EmpSeg = ""

   g_str_Parame = "SELECT * FROM MNT_SEGEMP WHERE "
   
   If p_TipSeg = 1 Then
      g_str_Parame = g_str_Parame & "SEGEMP_SEGPRE = 1 "
   ElseIf p_TipSeg = 2 Then
      g_str_Parame = g_str_Parame & "SEGEMP_SEGVIV = 1 "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   moddat_gf_Consulta_EmpSeg = Trim$(g_rst_Genera!SEGEMP_CODIGO)
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_ComSeg(ByVal p_Codigo As String) As String
   moddat_gf_Consulta_ComSeg = ""

   g_str_Parame = "SELECT * FROM MNT_SEGEMP WHERE SEGEMP_CODIGO = '" & p_Codigo & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   moddat_gf_Consulta_ComSeg = Trim$(g_rst_Listas!SEGEMP_RAZSOC)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_TipSeg(ByVal p_Codigo As String, ByVal p_TipSeg As Integer) As String
Dim r_str_Parame     As String

   moddat_gf_Consulta_TipSeg = ""

   r_str_Parame = "SELECT * FROM MNT_SEGTIP WHERE SEGTIP_CODIGO = '" & p_Codigo & "' AND SEGTIP_TIPSEG = " & CStr(p_TipSeg)

   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   moddat_gf_Consulta_TipSeg = Trim$(g_rst_Listas!SEGTIP_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_AplSeg(ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_Codigo As String, ByVal p_TipSeg As Integer, ByVal p_TipMon As Integer, ByVal p_MtoPre As Double, p_Combo As ComboBox, ByVal p_CodCiu As String) As Double
   moddat_gf_Consulta_AplSeg = 0

   If p_TipSeg > 0 Then    'Seguro de Préstamo
      g_str_Parame = "SELECT * FROM MNT_SEGPRE WHERE "
      g_str_Parame = g_str_Parame & "SEGPRE_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODIGO = '" & p_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPSEG = " & CStr(p_TipSeg) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMAX >= " & CStr(p_MtoPre) & " "
   Else                    'Seguro de Vivienda
      g_str_Parame = "SELECT * FROM MNT_SEGVIV WHERE "
      g_str_Parame = g_str_Parame & "SEGVIV_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODIGO = '" & p_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMAX >= " & CStr(p_MtoPre) & " "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   If p_TipSeg > 0 Then
      Call gs_BuscarCombo_Item(p_Combo, g_rst_Genera!SEGPRE_VTATIP)
      moddat_gf_Consulta_AplSeg = g_rst_Genera!SEGPRE_VTAFOI
   Else
      Call gs_BuscarCombo_Item(p_Combo, g_rst_Genera!SEGVIV_VTATIP)
      moddat_gf_Consulta_AplSeg = g_rst_Genera!SEGVIV_VTAFOI
   End If
   
   If Trim(p_CodCiu) = "7522" Or Trim(p_CodCiu) = "7523" Then
      Select Case p_TipSeg
         Case 11
            moddat_gf_Consulta_AplSeg = 0.068
         Case 12
            moddat_gf_Consulta_AplSeg = 0.1189
      End Select
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gf_Consulta_ValSeg(ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodEmp As String, ByVal p_TipSeg As Integer, ByVal p_TipMon As Integer, ByVal p_MtoPre As Double, ByRef p_TipVal As Integer, ByRef p_Import As Double, ByVal p_TasEsp As Integer)
Dim r_int_PosCad  As Integer
Dim r_dbl_TasEsp  As Double

   p_TipVal = 0
   p_Import = 0

   If p_TipSeg > 0 Then    'Seguro de Préstamo
      g_str_Parame = "SELECT * FROM MNT_SEGPRE WHERE "
      g_str_Parame = g_str_Parame & "SEGPRE_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODIGO = '" & p_CodEmp & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPSEG = " & CStr(p_TipSeg) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMAX >= " & CStr(p_MtoPre) & " "
   Else                    'Seguro de Vivienda
      g_str_Parame = "SELECT * FROM MNT_SEGVIV WHERE "
      g_str_Parame = g_str_Parame & "SEGVIV_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODIGO = '" & p_CodEmp & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMAX >= " & CStr(p_MtoPre) & " "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   If p_TipSeg > 0 Then
      p_TipVal = g_rst_Genera!SEGPRE_VTATIP
      p_Import = g_rst_Genera!SEGPRE_VTAFOI
   Else
      p_TipVal = g_rst_Genera!SEGVIV_VTATIP
      p_Import = g_rst_Genera!SEGVIV_VTAFOI
   End If
   
   'If Trim(p_CodCiu) = "7522" Or Trim(p_CodCiu) = "7523" Then
   '   Select Case p_TipSeg
   '      Case 11
   '         p_Import = 0.068
   '      Case 12
   '         p_Import = 0.1189
   '   End Select
   'End If
   
   If p_TipSeg > 0 Then
      If p_TasEsp > 1 Then
         r_int_PosCad = InStr(1, moddat_gf_Consulta_ParDes("522", CStr(p_TasEsp)), "%", vbTextCompare)
         r_dbl_TasEsp = Mid(moddat_gf_Consulta_ParDes("522", CStr(p_TasEsp)), 1, r_int_PosCad - 1)
         p_Import = p_Import + (p_Import * (r_dbl_TasEsp / 100))
      End If
      
'      If p_TasEsp = 1 Then
'         p_Import = p_Import + (p_Import * 0)
'      ElseIf p_TasEsp = 2 Then
'         p_Import = p_Import + (p_Import * (50 / 100))
'      ElseIf p_TasEsp = 3 Then
'         p_Import = p_Import + (p_Import * (100 / 100))
'      ElseIf p_TasEsp = 4 Then
'         p_Import = p_Import + (p_Import * (150 / 100))
'      ElseIf p_TasEsp = 5 Then
'         p_Import = p_Import + (p_Import * (200 / 100))
'      ElseIf p_TasEsp = 6 Then
'         p_Import = p_Import + (p_Import * (75 / 100))
'      End If
   End If
   
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Carter(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM MNT_CARTER WHERE CARTER_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY CARTER_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!CARTER_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!CARTER_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!CARTER_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = Trim$(g_rst_Genera!CARTER_TIPMON)
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = Trim$(g_rst_Genera!CARTER_IMPORT)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Busca_SecEco(p_CodGir As String) As String
   Dim r_rst_Genera     As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_GIRCOM WHERE GIRCOM_CODIGO = '" & p_CodGir & "'"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If
   
   r_rst_Genera.MoveFirst
   moddat_gf_Busca_SecEco = Trim(r_rst_Genera!GIRCOM_CODSEC)
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Function moddat_gf_Busca_GirCom(p_CodGir As String) As String
   moddat_gf_Busca_GirCom = ""
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_GIRCOM WHERE GIRCOM_CODIGO = '" & p_CodGir & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   moddat_gf_Busca_GirCom = Trim(g_rst_Genera!GIRCOM_DESCRI) & " (" & Trim(g_rst_Genera!GIRCOM_CODIGO) & ")"
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Consulta_DatInm(ByVal p_NumSol As String, ByRef p_Direcc As String, ByRef p_Distri As String, Optional ByRef p_CodPry As String, Optional ByRef p_NomPry As String, Optional ByRef p_CodBco As String, Optional ByVal p_FlgDis As Integer)
   Dim l_rst_Direcc  As ADODB.Recordset
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   p_Direcc = ""
   p_Distri = ""
   p_CodPry = ""
   p_NomPry = ""
   p_CodBco = ""
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & p_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Direcc, 3) Then
      Exit Sub
   End If
   
   If l_rst_Direcc.BOF And l_rst_Direcc.EOF Then
      l_rst_Direcc.Close
      Set l_rst_Direcc = Nothing
      Exit Sub
   End If
   
   l_rst_Direcc.MoveFirst
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(l_rst_Direcc!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(l_rst_Direcc!SOLINM_TIPZON))
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(l_rst_Direcc!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(l_rst_Direcc!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(l_rst_Direcc!SOLINM_UBIGEO))
   
   p_Direcc = r_str_TipVia & " " & Trim(l_rst_Direcc!SOLINM_NOMVIA) & " " & Trim(l_rst_Direcc!SOLINM_NUMVIA)

   If Len(Trim(Trim(l_rst_Direcc!SOLINM_INTDPT))) > 0 Then
      p_Direcc = p_Direcc & " - DPTO / INT.: " & Trim(l_rst_Direcc!SOLINM_INTDPT)
   End If

   If Len(Trim(Trim(l_rst_Direcc!SOLINM_NOMZON))) > 0 Then
      p_Direcc = p_Direcc & " - " & r_str_TipZon & " " & Trim(l_rst_Direcc!SOLINM_NOMZON)
   End If
      
   If p_FlgDis = 1 Then
      p_Distri = r_str_Distri
   Else
      p_Distri = r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   End If
   
   'If l_rst_Direcc!SOLINM_PRYMCS = 1 Then
      p_CodPry = Trim(l_rst_Direcc!SOLINM_PRYCOD & "")
   'ElseIf l_rst_Direcc!SOLINM_PRYMCS = 2 And Len(Trim(l_rst_Direcc!SOLINM_PRYNOM)) > 0 Then
      p_CodBco = Trim(l_rst_Direcc!SOLINM_PRYBCO & "")
      p_NomPry = Trim(l_rst_Direcc!SOLINM_PRYNOM & "")
   'End If
   
   l_rst_Direcc.Close
   Set l_rst_Direcc = Nothing
End Sub

Public Function moddat_gf_Consulta_AplSeg_Factor(ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_Codigo As String, ByVal p_TipSeg As Integer, ByVal p_TipMon As Integer, ByVal p_MtoPre As Double, ByRef p_TipApl As Integer) As Double
   moddat_gf_Consulta_AplSeg_Factor = 0

   If p_TipSeg > 0 Then    'Seguro de Préstamo
      g_str_Parame = "SELECT * FROM MNT_SEGPRE WHERE "
      g_str_Parame = g_str_Parame & "SEGPRE_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_CODIGO = '" & p_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPSEG = " & CStr(p_TipSeg) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGPRE_IMPMAX >= " & CStr(p_MtoPre) & " "
   Else                    'Seguro de Vivienda
      g_str_Parame = "SELECT * FROM MNT_SEGVIV WHERE "
      g_str_Parame = g_str_Parame & "SEGVIV_CODPRD = '" & p_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODSUB = '" & p_CodSub & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_CODIGO = '" & p_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "SEGVIV_TIPMON = " & CStr(p_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMIN <= " & CStr(p_MtoPre) & " AND "
      g_str_Parame = g_str_Parame & "SEGVIV_IMPMAX >= " & CStr(p_MtoPre) & " "
   End If

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   If p_TipSeg > 0 Then
      p_TipApl = g_rst_Genera!SEGPRE_VTATIP
      moddat_gf_Consulta_AplSeg_Factor = g_rst_Genera!SEGPRE_VTAFOI
   Else
      p_TipApl = g_rst_Genera!SEGVIV_VTATIP
      moddat_gf_Consulta_AplSeg_Factor = g_rst_Genera!SEGVIV_VTAFOI
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Carga_AtrGar(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipGar As String, ByVal p_TipMon As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM GARANTIA_ATRIBUTOS WHERE "
   g_str_Parame = g_str_Parame & "GARANTIA_TIPO = '" & p_TipGar & "' AND "
   g_str_Parame = g_str_Parame & "COD_MONEDA = '" & Format(p_TipMon, "000") & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DESCRIPCION ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem UCase(Trim$(g_rst_Genera!Descripcion))
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!GARANTIA_ATRIB)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_TipGar(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM GARANTIA_TIPO "
   g_str_Parame = g_str_Parame & "ORDER BY DESCRIPCION ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem UCase(Trim$(g_rst_Genera!Descripcion))
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!GARANTIA_TIPO)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_PreGar(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM AUDIT_GARANTIA_PREFERIDA "
   g_str_Parame = g_str_Parame & "ORDER BY DESCRIPCION ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem UCase(Trim$(g_rst_Genera!Descripcion))
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!GARANTIA_PREF)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_TipGar(ByVal p_TipGar As String) As String
   moddat_gf_Consulta_TipGar = ""

   g_str_Parame = "SELECT * FROM GARANTIA_TIPO WHERE GARANTIA_TIPO = '" & p_TipGar & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      moddat_gf_Consulta_TipGar = UCase(Trim$(g_rst_Genera!Descripcion))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_AtrGar(ByVal p_AtrGar As String) As String
   moddat_gf_Consulta_AtrGar = ""

   g_str_Parame = "SELECT * FROM GARANTIA_ATRIBUTOS WHERE GARANTIA_ATRIB = '" & p_AtrGar & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      moddat_gf_Consulta_AtrGar = UCase(Trim$(g_rst_Genera!Descripcion))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_PreGar(ByVal p_TipGar As String) As String
   moddat_gf_Consulta_PreGar = ""

   g_str_Parame = "SELECT * FROM GARANTIA_TIPO WHERE GARANTIA_TIPO = '" & p_TipGar & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      moddat_gf_Consulta_PreGar = UCase(Trim$(g_rst_Genera!GARANTIA_PREF))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Carga_UsuSis(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM SEG_USUMAE WHERE USUMAE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY USUMAE_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!USUMAE_CODIGO) & " - " & Trim(g_rst_Genera!USUMAE_NOMBRE)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!USUMAE_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!USUMAE_NOMBRE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Inicia_ActEco(ByVal p_TipCli As Integer, ByVal p_Indice As Integer)
   If p_TipCli = 1 Then
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_OrdAct = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_TipAct = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_SitTra = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = "000000"
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FreHab = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCar = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCar = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomAre = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumAnx = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TelDir = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Celula = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_DirEle = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipVia = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IntDpt = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipZon = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomZon = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_UbiGeo = "000000"
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Refere = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumFax = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCiu = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FlgEmp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCar = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCar = ""
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipDoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumDoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RazSoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCom = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipVia = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IntDpt = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipZon = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomZon = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_UbiGeo = "000000"
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Refere = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumFax = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCiu = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_GirCom = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_VtaMen = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCar = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCar = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RegTri = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_PorPar = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipLoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_AlqMen = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomArr = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TelArr = ""
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo = "000000"
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_PorPar = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt = ""
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele11 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele21 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe1 = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_SegPro = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele12 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele22 = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe2 = 0
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_IngNet = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Activi = ""
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_CodCiu = 0
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Observ = ""
   Else
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_OrdAct = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_TipAct = 0
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_SitTra = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = "000000"
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FreHab = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCar = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCar = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomAre = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumAnx = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TelDir = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Celula = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_DirEle = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant = ""
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipVia = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IntDpt = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipZon = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomZon = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_UbiGeo = "000000"
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Refere = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumFax = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCiu = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FlgEmp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCar = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCar = ""
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipDoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumDoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RazSoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCom = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipVia = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IntDpt = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipZon = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomZon = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_UbiGeo = "000000"
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Refere = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumFax = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCiu = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_GirCom = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_VtaMen = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCar = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCar = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RegTri = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_PorPar = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipLoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_AlqMen = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomArr = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TelArr = ""
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo = "000000"
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_PorPar = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt = ""
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele11 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele21 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe1 = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_SegPro = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele12 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele22 = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe2 = 0
   
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_IngNet = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Activi = ""
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_CodCiu = 0
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Observ = ""
   End If
End Sub

Public Sub moddat_gs_Inicia_DatCyg()
   moddat_g_arr_DatCyg(1).DatCli_TipDoc = 0
   moddat_g_arr_DatCyg(1).DatCli_NumDoc = ""
   moddat_g_arr_DatCyg(1).DatCli_DocAlt = 0
   moddat_g_arr_DatCyg(1).DatCli_TDoAlt = 0
   moddat_g_arr_DatCyg(1).DatCli_NDoAlt = ""
   moddat_g_arr_DatCyg(1).DatCli_ApePat = ""
   moddat_g_arr_DatCyg(1).DatCli_ApeMat = ""
   moddat_g_arr_DatCyg(1).DatCli_ApeCas = ""
   moddat_g_arr_DatCyg(1).DatCli_Nombre = ""
   moddat_g_arr_DatCyg(1).DatCli_FecNac = ""
   moddat_g_arr_DatCyg(1).DatCli_ApePat = ""
   moddat_g_arr_DatCyg(1).DatCli_Paises = ""
   moddat_g_arr_DatCyg(1).DatCli_UbiGeo = "000000"
   moddat_g_arr_DatCyg(1).DatCli_NivEst = 0
   moddat_g_arr_DatCyg(1).DatCli_Profes = ""
   moddat_g_arr_DatCyg(1).DatCli_Celula = ""
   moddat_g_arr_DatCyg(1).DatCli_DirEle = ""
   moddat_g_arr_DatCyg(1).DatCli_ChkEle = 0
   moddat_g_arr_DatCyg(1).DatCli_ClaSbs = ""
   moddat_g_arr_DatCyg(1).DatCli_ClasMC = ""
   moddat_g_arr_DatCyg(1).DatCli_ActEco = 0
End Sub

Public Sub moddat_gs_Carga_SubPrd(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String)
   ReDim p_Arregl(0)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM CRE_SUBPRD WHERE "
   g_str_Parame = g_str_Parame & "SUBPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SUBPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY SUBPRD_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!SUBPRD_CODSUB) & " - " & Trim(g_rst_Genera!SUBPRD_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!SUBPRD_CODSUB)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!SUBPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipMon = g_rst_Genera!SUBPRD_TIPMON
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!SUBPRD_MTOMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!SUBPRD_MTOMAX
      p_Arregl(UBound(p_Arregl)).Genera_PlzMin = g_rst_Genera!SUBPRD_PLZMIN
      p_Arregl(UBound(p_Arregl)).Genera_PlzMax = g_rst_Genera!SUBPRD_PLZMAX
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_SubPrd(p_CodPrd As String, ByVal p_CodSub As String) As String
   moddat_gf_Consulta_SubPrd = ""

   g_str_Parame = "SELECT * FROM CRE_SUBPRD WHERE "
   g_str_Parame = g_str_Parame & "SUBPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SUBPRD_CODSUB = '" & p_CodSub & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_SubPrd = Trim$(g_rst_Listas!SUBPRD_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_SubPrd_Arregl(p_Arregl() As moddat_tpo_Genera, p_CodPrd As String, ByVal p_CodSub As String) As Integer
   moddat_gf_Consulta_SubPrd_Arregl = False
   
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_SUBPRD WHERE "
   g_str_Parame = g_str_Parame & "SUBPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "SUBPRD_CODSUB = '" & p_CodSub & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_SubPrd_Arregl = True
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Listas!SUBPRD_CODSUB)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Listas!SUBPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipMon = g_rst_Listas!SUBPRD_TIPMON
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Listas!SUBPRD_MTOMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Listas!SUBPRD_MTOMAX
      p_Arregl(UBound(p_Arregl)).Genera_PlzMin = g_rst_Listas!SUBPRD_PLZMIN
      p_Arregl(UBound(p_Arregl)).Genera_PlzMax = g_rst_Listas!SUBPRD_PLZMAX
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_ParSubPrd_ComboItem(p_Combo As ComboBox, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARPRD_CODITE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParSubPrd(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParSubPrd_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String, Optional ByVal p_TipMon As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   If p_TipMon > 0 Then
      g_str_Parame = g_str_Parame & "SUBSTR(PARPRD_CODITE,3,1) = '" & CStr(p_TipMon) & "' AND "
   End If
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_ParSubPrd(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
   moddat_gf_Consulta_ParSubPrd = False
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
      
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_gf_Consulta_ParSubPrd = True
End Function

Public Function moddat_gf_Consulta_GirCom(p_CodGir As String) As String
   moddat_gf_Consulta_GirCom = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_GIRCOM WHERE GIRCOM_CODIGO = '" & p_CodGir & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_GirCom = Trim(g_rst_Listas!GIRCOM_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer) As Integer
   moddat_gf_Consulta_ActEco = 0
   
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_ActEco = g_rst_Listas!ACTECO_CODACT
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Buscar_NomMod(ByVal p_CodPrd As String, ByVal p_CodMod As String) As String
   moddat_gf_Buscar_NomMod = ""
   
   If moddat_gf_Consulta_ParPrd(moddat_g_arr_Genera(), p_CodPrd, "003", Format(p_CodMod, "000")) Then
      moddat_gf_Buscar_NomMod = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
   End If
End Function

Public Function moddat_gf_FecIng_Ins(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   moddat_gf_FecIng_Ins = ""
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
   g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGUIM_CODINS = " & CStr(p_CodIns)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_FecIng_Ins = gf_FormatoFecha(CStr(g_rst_Listas!SEGUIM_FECINI))
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Valida_Observ(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As Integer
   moddat_gf_Valida_Observ = False
   
   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      If g_rst_Listas!SEGFECACT = 0 Then
         moddat_gf_Valida_Observ = True
         Exit Do
      End If
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_EnvCor(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Arregl() As moddat_tpo_Genera, p_Asunto As String, p_Contenido As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
   Dim r_int_Contad      As Integer
   
   On Error GoTo moddat_gf_EnvCor

   'Inicializa
   p_Sesion.DownLoadMail = False
   p_Sesion.NewSession = True
   p_Sesion.SignOn
   p_Mensaje.SessionID = p_Sesion.SessionID
  
   'Envío
   p_Mensaje.Compose
   
   For r_int_Contad = 0 To UBound(p_Arregl) - 1
      If Len(Trim(p_Arregl(r_int_Contad + 1).Genera_Codigo)) > 0 Then
         p_Mensaje.RecipIndex = r_int_Contad
         p_Mensaje.RecipDisplayName = p_Arregl(r_int_Contad + 1).Genera_Codigo
      End If
   Next r_int_Contad

   p_Mensaje.MsgSubject = p_Asunto
   p_Mensaje.MsgNoteText = p_Contenido
   
   If Len(Trim(p_NomFil)) > 0 Then
      p_Mensaje.AttachmentIndex = 0
      p_Mensaje.AttachmentName = p_NomFil
      p_Mensaje.AttachmentPathName = p_RutFil & p_NomFil
      p_Mensaje.AttachmentPosition = 0
      p_Mensaje.AttachmentType = mapData
   End If
   
   p_Mensaje.Send
   DoEvents
  
  'Cierra la sesión
  p_Sesion.SignOff
  Exit Sub
  
moddat_gf_EnvCor:
   p_Sesion.SignOff
   MsgBox Err.Description, vbCritical
End Sub

Public Function moddat_gf_Buscar_DirEle_Codigo(ByVal p_CodEje As String) As String
   moddat_gf_Buscar_DirEle_Codigo = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODEJE = '" & p_CodEje & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Buscar_DirEle_Codigo = Trim(g_rst_Genera!EJECMC_DIRELE & "")
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Buscar_DirEle_UsuSis(ByVal p_UsuSis As String) As String
   moddat_gf_Buscar_DirEle_UsuSis = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODUSU = '" & p_UsuSis & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Buscar_DirEle_UsuSis = Trim(g_rst_Genera!EJECMC_DIRELE & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Buscar_CodEje_UsuSis(ByVal p_UsuSis As String) As String
   'Devuelve el Código de Ejecutivo miCasita recibiendo el Código de Acceso al Sistema
   moddat_gf_Buscar_CodEje_UsuSis = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC WHERE "
   g_str_Parame = g_str_Parame & "EJECMC_CODUSU = '" & p_UsuSis & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Buscar_CodEje_UsuSis = Trim(g_rst_Genera!EJECMC_CODEJE & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Buscar_DirEle_TipUsu(ByVal p_TipUsu As Integer) As String
   moddat_gf_Buscar_DirEle_TipUsu = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJECMC A, CRE_EJETIP B WHERE "
   g_str_Parame = g_str_Parame & "A.EJECMC_CODEJE = B.EJETIP_CODEJE AND "
   g_str_Parame = g_str_Parame & "A.EJECMC_SITUAC <> 2 AND "
   g_str_Parame = g_str_Parame & "B.EJETIP_TIPEJE = " & CStr(p_TipUsu) & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Buscar_DirEle_TipUsu = Trim(g_rst_Genera!EJECMC_DIRELE & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Verifica_DirEle(p_Arregl() As moddat_tpo_Genera, ByVal p_Codigo As String) As Integer
   Dim r_int_Contad     As Integer
   
   moddat_gf_Verifica_DirEle = False
   
   If Len(Trim(p_Codigo)) = 0 Then
      moddat_gf_Verifica_DirEle = True
      Exit Function
   End If
   
   For r_int_Contad = 1 To UBound(p_Arregl)
      If p_Arregl(r_int_Contad).Genera_Codigo = p_Codigo Then
         moddat_gf_Verifica_DirEle = True
         Exit For
      End If
   Next r_int_Contad
End Function

Public Function moddat_gf_UsuObs(ByVal p_NumSol As String, ByVal p_CodIns As Integer) As String
   moddat_gf_UsuObs = ""

   g_str_Parame = "SELECT * FROM TRA_SEGDET WHERE "
   g_str_Parame = g_str_Parame & "SEGDET_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODINS = " & CStr(p_CodIns) & " AND "
   g_str_Parame = g_str_Parame & "SEGDET_CODOCU = 21 "
   g_str_Parame = g_str_Parame & "ORDER BY SEGDET_NUMOBS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      moddat_gf_UsuObs = Trim(g_rst_Princi!SEGUSUCRE & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Public Function moddat_gf_ComMVi(ByVal p_CodPrd As String, ByVal p_TipCom As Integer, ByVal p_TipMon As Integer, ByVal p_PlaAno As Integer) As Double
   moddat_gf_ComMVi = 0

   g_str_Parame = "SELECT * FROM OPE_COMMVI WHERE "
   g_str_Parame = g_str_Parame & "COMMVI_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "COMMVI_TIPCOM = " & CStr(p_TipCom) & " AND "
   g_str_Parame = g_str_Parame & "COMMVI_TIPMON = " & CStr(p_TipMon) & " AND "
   g_str_Parame = g_str_Parame & "COMMVI_PLAINI <= " & CStr(p_PlaAno) & " AND "
   g_str_Parame = g_str_Parame & "COMMVI_PLAFIN >= " & CStr(p_PlaAno) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_ComMVi = g_rst_Listas!COMMVI_PORCEN
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function gf_Formato_NumSol(ByVal p_NumSol As String) As String
   gf_Formato_NumSol = Left(p_NumSol, 3) & "-" & Mid(p_NumSol, 4, 3) & "-" & Mid(p_NumSol, 7, 2) & "-" & Right(p_NumSol, 4)
End Function

Public Function gf_Formato_NumOpe(ByVal p_NumOpe As String) As String
   gf_Formato_NumOpe = Left(p_NumOpe, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Right(p_NumOpe, 5)
End Function

Public Function gf_Formato_NumRef(ByVal p_NumRef As String, ByVal p_Tipo As Integer) As String
   p_NumRef = Format(p_NumRef, "0000000000")
   
   If p_Tipo = 0 Then
      gf_Formato_NumRef = Left(p_NumRef, 4) & "-" & Mid(p_NumRef, 5, 2) & "-" & Right(p_NumRef, 4)
   Else
      gf_Formato_NumRef = Mid(p_NumRef, 1, 1) & Mid(p_NumRef, 2, 2) & "-" & Mid(p_NumRef, 4, 2) & "-" & Right(p_NumRef, 5)
   End If
End Function

Public Function gf_Pago_GasAdm(ByVal p_NumSol As String) As String
   gf_Pago_GasAdm = "NO"
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & p_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      gf_Pago_GasAdm = "SI"
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_PerCon(p_Combo As ComboBox, ByVal p_CodPer As String, ByVal p_TipTab As Integer)
   p_Combo.Clear

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * "
   g_str_Parame = g_str_Parame & "   FROM MNT_PERCON "
   g_str_Parame = g_str_Parame & "  WHERE PERCON_CODEMP = '" & p_CodPer & "'"
   g_str_Parame = g_str_Parame & "    AND PERCON_TIPTAB = " & p_TipTab
   g_str_Parame = g_str_Parame & "    AND PERCON_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY PERCON_NOMBRE ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!PERCON_NOMBRE)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!PERCON_CODCON)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_PerCon_OLD(p_Combo As ComboBox, ByVal p_CodPer As String)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PERCON WHERE "
   g_str_Parame = g_str_Parame & "PERCON_CODEMP = '" & p_CodPer & "' AND "
   g_str_Parame = g_str_Parame & "PERCON_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PERCON_NOMBRE ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!PERCON_NOMBRE)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!PERCON_CODCON)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_PerCon(ByVal p_CodEmp As String, ByVal p_CodCon As String, ByVal p_TipTab As Integer, Optional ByRef p_DirEle As String) As String
   moddat_gf_Consulta_PerCon = ""
   p_DirEle = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM MNT_PERCON "
   g_str_Parame = g_str_Parame & " WHERE PERCON_CODEMP = '" & p_CodEmp & "'"
   g_str_Parame = g_str_Parame & "   AND PERCON_CODCON = '" & p_CodCon & "' "
   g_str_Parame = g_str_Parame & "   AND PERCON_TIPTAB = " & p_TipTab
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_PerCon = Trim(g_rst_Listas!PERCON_NOMBRE)
      p_DirEle = Trim(g_rst_Listas!PERCON_DIRELE & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_PerCon_OLD(ByVal p_CodEmp As String, ByVal p_CodCon As String, Optional ByRef p_DirEle As String) As String
   moddat_gf_Consulta_PerCon_OLD = ""
   p_DirEle = ""
   
   g_str_Parame = "SELECT * FROM MNT_PERCON WHERE "
   g_str_Parame = g_str_Parame & "PERCON_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "PERCON_CODCON = '" & p_CodCon & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_PerCon_OLD = Trim(g_rst_Listas!PERCON_NOMBRE)
      p_DirEle = Trim(g_rst_Listas!PERCON_DIRELE & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_PerTas(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodEmp As String)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM MNT_PERPER WHERE "
   g_str_Parame = g_str_Parame & "PERPER_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "PERPER_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PERPER_NOMPER ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!PERPER_NOMPER)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!PERPER_CODPER)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!PERPER_NOMPER)
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = Trim(g_rst_Listas!PERPER_CODREP)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_ProMas(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_FlgSit As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM SEG_BATMAE "
   If p_FlgSit = 1 Then
      g_str_Parame = g_str_Parame & "WHERE BATMAE_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY BATMAE_CODPRG ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!BATMAE_CODPRG) & " - " & Trim$(g_rst_Listas!BATMAE_NOMPRG)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!BATMAE_CODPRG)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!BATMAE_NOMPRG)
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = Trim(g_rst_Listas!BATMAE_CODFRE)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Buscar_CodEje_Cargos(ByVal p_CodEje As String) As String
   Dim r_rst_Genera     As ADODB.Recordset
   
   moddat_gf_Buscar_CodEje_Cargos = ""
   
   g_str_Parame = "SELECT * FROM CRE_EJETIP WHERE "
   g_str_Parame = g_str_Parame & "EJETIP_CODEJE = '" & p_CodEje & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      
      Do While Not r_rst_Genera.EOF
         moddat_gf_Buscar_CodEje_Cargos = moddat_gf_Buscar_CodEje_Cargos & moddat_gf_Consulta_ParDes("034", CStr(r_rst_Genera!EJETIP_TIPEJE)) & " - "
         r_rst_Genera.MoveNext
      Loop
      
      moddat_gf_Buscar_CodEje_Cargos = Mid(moddat_gf_Buscar_CodEje_Cargos, 1, Len(moddat_gf_Buscar_CodEje_Cargos) - 3)
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Carga_Usuario(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipEje As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_EJECMC "
   
   If p_TipEje = 1 Then
      g_str_Parame = g_str_Parame & "WHERE EJECMC_SITUAC = 1 "
   End If
   
   g_str_Parame = g_str_Parame & "ORDER BY EJECMC_APEPAT ASC, EJECMC_APEMAT ASC, EJECMC_NOMBRE ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!EJECMC_APEPAT) & " " & Trim$(g_rst_Listas!EJECMC_APEMAT) & " " & Trim$(g_rst_Listas!EJECMC_NOMBRE)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!EJECMC_CODEJE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!EJECMC_APEPAT) & " " & Trim(g_rst_Listas!EJECMC_APEMAT) & " " & Trim(g_rst_Listas!EJECMC_NOMBRE)
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = Trim(g_rst_Listas!EJECMC_DIRELE)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_UsuTec(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_EJECMC A, CRE_EJETIP B "
   g_str_Parame = g_str_Parame & "WHERE EJECMC_SITUAC = 1 AND EJECMC_CODEJE = EJETIP_CODEJE AND EJETIP_TIPEJE >= 810 AND EJETIP_TIPEJE <= 819 "
   g_str_Parame = g_str_Parame & "ORDER BY EJECMC_APEPAT ASC, EJECMC_APEMAT ASC, EJECMC_NOMBRE ASC"
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!EJECMC_APEPAT) & " " & Trim$(g_rst_Listas!EJECMC_APEMAT) & " " & Trim$(g_rst_Listas!EJECMC_NOMBRE)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!EJECMC_CODEJE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!EJECMC_APEPAT) & " " & Trim(g_rst_Listas!EJECMC_APEMAT) & " " & Trim(g_rst_Listas!EJECMC_NOMBRE)
      p_Arregl(UBound(p_Arregl)).Genera_Prefij = Trim(g_rst_Listas!EJECMC_DIRELE)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_UsuCar(p_Lista As ListBox, ByVal p_CodEje As String)
   Dim r_rst_Genera     As ADODB.Recordset
   
   p_Lista.Clear
   
   g_str_Parame = "SELECT * FROM CRE_EJETIP WHERE "
   g_str_Parame = g_str_Parame & "EJETIP_CODEJE = '" & p_CodEje & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Sub
   End If
   
   r_rst_Genera.MoveFirst
   Do While Not r_rst_Genera.EOF
      p_Lista.AddItem moddat_gf_Consulta_ParDes("034", CStr(r_rst_Genera!EJETIP_TIPEJE))
      r_rst_Genera.MoveNext
   Loop
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_MenPlt(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPlt As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM SEG_PLTOPC WHERE "
   g_str_Parame = g_str_Parame & "PLTOPC_CODPLT = '" & p_CodPlt & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_FLGMEN = 1 AND "
   g_str_Parame = g_str_Parame & "PLTOPC_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PLTOPC_CODMEN ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!PLTOPC_CODMEN) & " - " & Trim(g_rst_Listas!PLTOPC_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!PLTOPC_CODMEN)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!PLTOPC_DESCRI)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_SubPlt(p_Combo As ComboBox, ByVal p_CodPlt As String, ByVal p_CodMen As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM SEG_PLTOPC WHERE "
   g_str_Parame = g_str_Parame & "PLTOPC_CODPLT = '" & p_CodPlt & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_CODMEN = '" & p_CodMen & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_FLGMEN = 2 AND "
   g_str_Parame = g_str_Parame & "PLTOPC_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PLTOPC_CODSUB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!PLTOPC_CODSUB) & " - " & Trim(g_rst_Listas!PLTOPC_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!PLTOPC_CODSUB)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_NomOpcPlt(ByVal p_CodPlt As String, ByVal p_CodMen As String, ByVal p_CodSub As String) As String
   moddat_gf_Consulta_NomOpcPlt = ""
   
   g_str_Parame = "SELECT * FROM SEG_PLTOPC WHERE "
   g_str_Parame = g_str_Parame & "PLTOPC_CODPLT = '" & p_CodPlt & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_CODMEN = '" & p_CodMen & "' AND "
   g_str_Parame = g_str_Parame & "PLTOPC_CODSUB = '" & p_CodSub & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_NomOpcPlt = Trim(g_rst_Listas!PLTOPC_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_Produc_Comerc(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipCre As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)

   g_str_Parame = "SELECT * FROM CRE_PRODUC WHERE PRODUC_SITCOM = 1 "
   If p_TipCre <> 99 Then
      g_str_Parame = g_str_Parame & "AND PRODUC_CODCLA = " & CStr(p_TipCre) & " "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY PRODUC_CODIGO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim$(g_rst_Listas!PRODUC_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Listas!Produc_Codigo)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Listas!PRODUC_DESCRI)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_ConsultaNombreBat(ByVal p_CodPro As String) As String
   moddat_gf_ConsultaNombreBat = ""

   g_str_Parame = "SELECT * FROM SEG_BATMAE WHERE BATMAE_CODPRG = '" & p_CodPro & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_ConsultaNombreBat = Trim(g_rst_Listas!BATMAE_NOMPRG)
   End If
End Function

Public Function moddat_gf_ObtieneTipCamDia(ByVal p_TipCam As Integer, ByVal p_TipMon As Integer, ByVal p_FecDia As String, ByVal p_TipTip As Integer) As Double
   moddat_gf_ObtieneTipCamDia = 0
   
   'Obteniendo Tipo de Cambio del Dia
   g_str_Parame = "SELECT * FROM OPE_TIPCAM WHERE "
   g_str_Parame = g_str_Parame & "TIPCAM_CODIGO = " & CStr(p_TipCam) & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_FECDIA = " & p_FecDia & " AND "
   g_str_Parame = g_str_Parame & "TIPCAM_TIPMON = " & CStr(p_TipMon) & " ORDER BY "
   g_str_Parame = g_str_Parame & "TIPCAM_HORDIA DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
   
      If p_TipTip = 1 Then
         moddat_gf_ObtieneTipCamDia = g_rst_Genera!TIPCAM_VENTAS
      ElseIf p_TipTip = 2 Then
         moddat_gf_ObtieneTipCamDia = g_rst_Genera!TIPCAM_COMPRA
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Sub moddat_gs_Carga_EmpGrp(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_EMPGRP WHERE "
   g_str_Parame = g_str_Parame & "EMPGRP_SITUAC <> 9 "
   g_str_Parame = g_str_Parame & "ORDER BY EMPGRP_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!EMPGRP_NOMCOM)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!EMPGRP_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!EMPGRP_RAZSOC)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_ConsultaSucAge(ByVal p_CodEmp As String, ByVal p_CodSuc As String) As String
   moddat_gf_ConsultaSucAge = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_SUCAGE WHERE "
   g_str_Parame = g_str_Parame & "SUCAGE_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "SUCAGE_CODSUC = '" & p_CodSuc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Function
   End If
   
   g_rst_Listas.MoveFirst
   moddat_gf_ConsultaSucAge = Trim(g_rst_Listas!SUCAGE_DESCRI)
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_ConsultaEmpGrp(p_CodEmp As String) As String
   moddat_gf_ConsultaEmpGrp = ""

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_EMPGRP WHERE "
   g_str_Parame = g_str_Parame & "EMPGRP_CODIGO = '" & p_CodEmp & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_ConsultaEmpGrp = Trim(g_rst_Listas!EMPGRP_RAZSOC)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_LisIte(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_TipTab As Integer, ByVal p_CodGrp As String, Optional ByVal p_TipOrd As Integer)
   'p_TipTab   =  1  -  MNT$PARDES - Parámetros de Descripción
   'p_TipTab   =  2  -  MNT$PARVAL - Parámetros de Valor
   
   p_Combo.Clear
   ReDim p_Arregl(0)

   Select Case p_TipTab
      Case 1
         g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
         g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & p_CodGrp & "' AND "
         g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
         g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
         
         If p_TipOrd = 1 Then
            g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"
         Else
            g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"
         End If
      
      Case 2
         g_str_Parame = "SELECT * FROM MNT_PARVAL WHERE "
         g_str_Parame = g_str_Parame & "PARVAL_CODGRP = '" & p_CodGrp & "' AND "
         g_str_Parame = g_str_Parame & "PARVAL_CODITE <> '000000' AND "
         g_str_Parame = g_str_Parame & "PARVAL_SITUAC = 1 "
         g_str_Parame = g_str_Parame & "ORDER BY PARVAL_DESCRI ASC"
   End Select

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      Select Case p_TipTab
         Case 1:     p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
         Case 2:     p_Combo.AddItem Trim$(g_rst_Genera!PARVAL_DESCRI)
      End Select
         
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      Select Case p_TipTab
         Case 1
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARDES_CODITE)
            p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARDES_DESCRI)
            p_Arregl(UBound(p_Arregl)).Genera_TipVal = 0
            p_Arregl(UBound(p_Arregl)).Genera_Cantid = 0
         Case 2
            p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARVAL_CODITE)
            p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARVAL_DESCRI)
            p_Arregl(UBound(p_Arregl)).Genera_TipVal = Trim$(g_rst_Genera!PARVAL_TIPVAL)
            p_Arregl(UBound(p_Arregl)).Genera_Cantid = Trim$(g_rst_Genera!PARVAL_CANTID)
      End Select
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_MotRec(p_Combo As ComboBox, ByVal p_CodIns As String)
   p_Combo.Clear

   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '003' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "(SUBSTR(PARDES_CODITE,4,1) = '0' OR "
   g_str_Parame = g_str_Parame & " SUBSTR(PARDES_CODITE,4,2) = '" & Format(p_CodIns, "00") & "') AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARDES_CODITE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Depart(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '101' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,3,2) = '00' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,5,2) = '00' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(Left(Trim(g_rst_Genera!PARDES_CODITE), 2))
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Provin(p_Combo As ComboBox, ByVal p_CodDpt As String)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '101' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,1,2) = '" & p_CodDpt & "' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,3,2) <> '00' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,5,2) = '00' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(Mid(Trim(g_rst_Genera!PARDES_CODITE), 3, 2))
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Distri(p_Combo As ComboBox, ByVal p_CodDpt As String, ByVal p_CodPrv As String)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '101' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,1,2) = '" & p_CodDpt & "' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,3,2) = '" & p_CodPrv & "' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,5,2) <> '00' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(Mid(Trim(g_rst_Genera!PARDES_CODITE), 5, 2))
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_Distri_Lima(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '101' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "(SUBSTR(PARDES_CODITE,1,4) = '0701' OR SUBSTR(PARDES_CODITE,1,4) = '1501') AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_CODITE,5,2) <> '00' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = Trim(g_rst_Genera!PARDES_CODITE)
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_DstZon(p_Combo As ComboBox, ByVal p_CodUbi As String)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_REFVIV WHERE "
   g_str_Parame = g_str_Parame & "SUBSTR(REFVIV_CODZON,1,6) = '" & p_CodUbi & "' "
   g_str_Parame = g_str_Parame & "ORDER BY REFVIV_DESCRI ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!REFVIV_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(Mid(Trim(g_rst_Genera!REFVIV_CODZON), 7, 2))
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_CdCIIU(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '102' AND "
   g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Right(Trim$(g_rst_Genera!PARDES_CODITE), 4) & " - " & Trim$(g_rst_Genera!PARDES_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(Trim(g_rst_Genera!PARDES_CODITE))
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_LisIte_Combo(p_Combo As ComboBox, ByVal p_TipTab As Integer, ByVal p_CodGrp As String, Optional ByVal p_TipOrd As Integer)
   'p_TipTab   =  1  -  MNT$PARDES - Parámetros de Descripción
   'p_TipTab   =  2  -  MNT$PARVAL - Parámetros de Valor
   
   p_Combo.Clear

   Select Case p_TipTab
      Case 1
         g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
         g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & p_CodGrp & "' AND "
         g_str_Parame = g_str_Parame & "PARDES_CODITE <> '000000' AND "
         g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
         If p_TipOrd = 1 Then
            g_str_Parame = g_str_Parame & "ORDER BY PARDES_DESCRI ASC"
         Else
            g_str_Parame = g_str_Parame & "ORDER BY PARDES_CODITE ASC"
         End If
      
      Case 2
         g_str_Parame = "SELECT * FROM MNT_PARVAL WHERE "
         g_str_Parame = g_str_Parame & "PARVAL_CODGRP = '" & p_CodGrp & "' AND "
         g_str_Parame = g_str_Parame & "PARVAL_CODITE <> '000000' AND "
         g_str_Parame = g_str_Parame & "PARVAL_SITUAC = 1 "
         If p_TipOrd = 1 Then
            g_str_Parame = g_str_Parame & "ORDER BY PARVAL_DESCRI ASC"
         Else
            g_str_Parame = g_str_Parame & "ORDER BY PARVAL_CODITE ASC"
         End If
   End Select

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      Select Case p_TipTab
         Case 1
            p_Combo.AddItem Trim$(g_rst_Genera!PARDES_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!PARDES_CODITE)
         Case 2
            p_Combo.AddItem Trim$(g_rst_Genera!PARVAL_DESCRI)
            p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!PARVAL_CODITE)
      End Select
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParPrd(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodGrp As String)
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParPrd_Combo(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodGrp As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParPrd_ComboItem(p_Combo As ComboBox, ByVal p_CodPrd As String, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!PARPRD_CODITE)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_ParAct(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodAct As String, ByVal p_CodGrp As String)
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = '" & p_CodAct & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE <> '000' AND "
   g_str_Parame = g_str_Parame & "PARACT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY PARACT_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARACT_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARACT_DESCRI)
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARACT_TIPVAL
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARACT_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARACT_CANTID
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARACT_VALMIN
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARACT_VALMAX
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Carga_TipCre(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCRE  "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCRE_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!TIPCRE_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!TIPCRE_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!TIPCRE_DESCRI)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_ClfCre(ByVal p_ClaCre As String, p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCLA  "
   g_str_Parame = g_str_Parame & "WHERE TIPCLA_TIPCRE = '" & p_ClaCre & "' "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCLA_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!TIPCLA_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!TIPCLA_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!TIPCLA_DESCRI)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_TipGar_Ctb(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPGAR  "
   g_str_Parame = g_str_Parame & "ORDER BY TIPGAR_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!TIPGAR_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!TIPGAR_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!TIPGAR_DESCRI)
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_ClaGar(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CLAGAR  "
   g_str_Parame = g_str_Parame & "ORDER BY CLAGAR_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!CLAGAR_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!CLAGAR_CODIGO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!CLAGAR_DESCRI)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_SucAge(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodEmp As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_SUCAGE WHERE "
   g_str_Parame = g_str_Parame & "SUCAGE_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "SUCAGE_SITUAC <> 9 "
   g_str_Parame = g_str_Parame & "ORDER BY SUCAGE_CODSUC ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!SUCAGE_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!SUCAGE_CODSUC)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!SUCAGE_DESCRI)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_ClaCtaCtb(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CLACTA  "
   g_str_Parame = g_str_Parame & "ORDER BY CLACTA_CODCLA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem CStr(CInt(g_rst_Listas!CLACTA_CODCLA)) & " - " & Trim(g_rst_Listas!CLACTA_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!CLACTA_CODCLA)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_TipMonCtb(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPMON  "
   g_str_Parame = g_str_Parame & "ORDER BY TIPMON_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem CStr(CInt(g_rst_Listas!TIPMON_CODIGO)) & " - " & Trim(g_rst_Listas!TIPMON_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!TIPMON_CODIGO)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_LibCtb(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_LIBCON  "
   g_str_Parame = g_str_Parame & "ORDER BY LIBCON_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem CStr(CInt(g_rst_Listas!LIBCON_CODIGO)) & " - " & Trim(g_rst_Listas!LIBCON_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!LIBCON_CODIGO)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_LibCtb(p_CodLib As Integer) As String
   moddat_gf_Consulta_LibCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_LIBCON  "
   g_str_Parame = g_str_Parame & "WHERE LIBCON_CODIGO = " & CStr(p_CodLib) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_LibCtb = Trim(g_rst_Listas!LIBCON_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_ParEmp(ByVal p_CodEmp As String, ByVal p_CodGrp As String, p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PAREMP  WHERE "
   g_str_Parame = g_str_Parame & "PAREMP_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "PAREMP_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PAREMP_CODITE <> '000000' "
   g_str_Parame = g_str_Parame & "ORDER BY PAREMP_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!PAREMP_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!PAREMP_CODITE)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!PAREMP_DESCRI)
      
      p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Listas!PAREMP_TIPPAR
      p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Listas!PAREMP_TIPVAL
      
      p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Listas!PAREMP_VALIND
      p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Listas!PAREMP_VALINI
      p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Listas!PAREMP_VALFIN
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_CtaCtb(ByVal p_CodEmp As String, p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera, ByVal p_CodCla As Integer, ByVal p_CodNiv As Integer, ByVal p_CodMon As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAMAE  WHERE "
   g_str_Parame = g_str_Parame & "CTAMAE_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "CTAMAE_REGCOM = 1 "
   
   If p_CodCla <> 0 Then
      g_str_Parame = g_str_Parame & "AND SUBSTR(CTAMAE_CODCTA, 1, 1) = '" & CStr(p_CodCla) & "' "
   End If
   
   If p_CodNiv <> 0 Then
      g_str_Parame = g_str_Parame & "AND CTAMAE_CODNIV = " & CStr(p_CodNiv) & " "
   End If
   
   If p_CodMon = -1 Then
      g_str_Parame = g_str_Parame & "AND SUBSTR(CTAMAE_CODCTA, 3, 1) <> '0' "
   Else
      g_str_Parame = g_str_Parame & "AND SUBSTR(CTAMAE_CODCTA, 3, 1) = '" & CStr(p_CodMon) & " ' "
   End If
   
   
   g_str_Parame = g_str_Parame & "ORDER BY CTAMAE_CODCTA ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!CTAMAE_CODCTA) & " - " & Trim(g_rst_Listas!CTAMAE_DESCRI)
      
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!CTAMAE_CODCTA)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!CTAMAE_DESCRI & "")
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_ParDes(ByVal p_CodGrp As String, ByVal p_CodIte As String) As String
Dim r_str_Parame     As String

   moddat_gf_Consulta_ParDes = ""
   
   p_CodIte = Format(p_CodIte, "000000")
   
   r_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   r_str_Parame = r_str_Parame & "PARDES_CODGRP = '" & p_CodGrp & "' AND "
   r_str_Parame = r_str_Parame & "PARDES_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_ParDes = Trim(g_rst_Listas!PARDES_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_ParVal(ByVal p_CodGrp As String, ByVal p_CodIte As String) As Double
   moddat_gf_Consulta_ParVal = 0
   
   p_CodIte = Format(p_CodIte, "000")
   
   g_str_Parame = "SELECT * FROM MNT_PARVAL WHERE "
   g_str_Parame = g_str_Parame & "PARVAL_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARVAL_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      moddat_gf_Consulta_ParVal = g_rst_Genera!PARVAL_CANTID
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_ParDes_1(ByVal p_CodGrp As String, ByVal p_CodIte As String, ByVal p_NumPos As Integer) As String
   moddat_gf_Consulta_ParDes_1 = ""
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(PARDES_DESCRI,1," & CStr(p_NumPos) & ") = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      moddat_gf_Consulta_ParDes_1 = Trim(g_rst_Listas!PARDES_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_ParPrd(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
   moddat_gf_Consulta_ParPrd = False
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARPRD WHERE "
   g_str_Parame = g_str_Parame & "PARPRD_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARPRD_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Function
   End If
   
   g_rst_Genera.MoveFirst
      
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARPRD_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARPRD_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARPRD_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARPRD_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARPRD_CANTID
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARPRD_VALMIN
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARPRD_VALMAX
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_gf_Consulta_ParPrd = True
End Function

Public Function moddat_gf_Consulta_ParAct(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_CodAct As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
   moddat_gf_Consulta_ParAct = False
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_PARACT WHERE "
   g_str_Parame = g_str_Parame & "PARACT_CODPRD = '" & p_CodPrd & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODSUB = '" & p_CodSub & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODACT = " & p_CodAct & " AND "
   g_str_Parame = g_str_Parame & "PARACT_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PARACT_CODITE = '" & p_CodIte & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Function
   End If
   
   g_rst_Genera.MoveFirst
      
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Genera!PARACT_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Genera!PARACT_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Genera!PARACT_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Genera!PARACT_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Genera!PARACT_CANTID
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Genera!PARACT_VALMIN
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Genera!PARACT_VALMAX
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_gf_Consulta_ParAct = True
End Function

Public Function moddat_gf_ConsultaPrefijClaGar(ByVal p_CodCla As String) As String
   moddat_gf_ConsultaPrefijClaGar = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CLAGAR  "
   g_str_Parame = g_str_Parame & "WHERE CLAGAR_CODIGO = '" & p_CodCla & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_ConsultaPrefijClaGar = Trim(g_rst_Listas!CLAGAR_PREFIJ)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_ConsultaClaseGar(ByVal p_CodCla As String) As String
   moddat_gf_ConsultaClaseGar = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CLAGAR  "
   g_str_Parame = g_str_Parame & "WHERE CLAGAR_CODIGO = '" & p_CodCla & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_ConsultaClaseGar = Trim(g_rst_Listas!CLAGAR_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_ConsultaClasifCred(ByVal p_ClaCre As String, ByVal p_ClfCre As String) As String
   moddat_gf_ConsultaClasifCred = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCLA  WHERE "
   g_str_Parame = g_str_Parame & "TIPCLA_TIPCRE = '" & p_ClaCre & "' AND "
   g_str_Parame = g_str_Parame & "TIPCLA_CODIGO = '" & p_ClfCre & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_ConsultaClasifCred = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_ConsultaClasifCredGral(ByVal p_ClfCre As String) As String
   moddat_gf_ConsultaClasifCredGral = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCLA  WHERE "
   g_str_Parame = g_str_Parame & "TIPCLA_CODIGO = '" & p_ClfCre & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_ConsultaClasifCredGral = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_ParEmp(p_Arregl() As moddat_tpo_Genera, ByVal p_CodEmp As String, ByVal p_CodGrp As String, ByVal p_CodIte As String) As Integer
   moddat_gf_Consulta_ParEmp = False
   
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PAREMP WHERE "
   g_str_Parame = g_str_Parame & "PAREMP_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "PAREMP_CODGRP = '" & p_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PAREMP_CODITE = '" & Format(p_CodIte, "000000") & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     
     Exit Function
   End If
   
   g_rst_Listas.MoveFirst
      
   ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
   
   p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!PAREMP_CODITE)
   p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!PAREMP_DESCRI)
   p_Arregl(UBound(p_Arregl)).Genera_TipPar = g_rst_Listas!PAREMP_TIPPAR
   p_Arregl(UBound(p_Arregl)).Genera_TipVal = g_rst_Listas!PAREMP_TIPVAL
   p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Listas!PAREMP_VALIND
   p_Arregl(UBound(p_Arregl)).Genera_ValMin = g_rst_Listas!PAREMP_VALINI
   p_Arregl(UBound(p_Arregl)).Genera_ValMax = g_rst_Listas!PAREMP_VALFIN
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   moddat_gf_Consulta_ParEmp = True
End Function

Public Function moddat_gf_ConsultaPerMesActivo(ByVal p_CodEmp As String, ByVal p_TipPer As Integer, Optional ByRef p_FecIni As String, Optional ByRef p_FecFin As String, Optional ByRef p_PerMes As Integer, Optional ByRef p_PerAno As Integer) As String
   Dim r_rst_Genera     As ADODB.Recordset

   moddat_gf_ConsultaPerMesActivo = ""
   p_FecIni = ""
   p_FecFin = ""
   p_PerMes = 0
   p_PerAno = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_PERMES WHERE "
   g_str_Parame = g_str_Parame & "PERMES_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & CStr(p_TipPer) & " AND "
   g_str_Parame = g_str_Parame & "PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      p_FecIni = gf_FormatoFecha(CStr(r_rst_Genera!PERMES_FECINI))
      p_FecFin = gf_FormatoFecha(CStr(r_rst_Genera!PERMES_FECFIN))
      p_PerMes = r_rst_Genera!PERMES_CODMES
      p_PerAno = r_rst_Genera!PERMES_CODANO
      
      moddat_gf_ConsultaPerMesActivo = moddat_gf_Consulta_ParDes("033", CStr(r_rst_Genera!PERMES_CODMES)) & " " & Format(r_rst_Genera!PERMES_CODANO, "0000")
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Public Function moddat_gf_Consulta_NomCtaCtb(ByVal p_CodEmp As String, ByVal p_CodCta As String) As String
   moddat_gf_Consulta_NomCtaCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAMAE  WHERE "
   g_str_Parame = g_str_Parame & "CTAMAE_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "CTAMAE_CODCTA = '" & p_CodCta & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_NomCtaCtb = Trim(g_rst_Listas!CTAMAE_DESCRI & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_NomEntFin(ByVal p_CodEmp As String) As String
   moddat_gf_Consulta_NomEntFin = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_EMPSUP  WHERE "
   g_str_Parame = g_str_Parame & "EMPSUP_CODIGO = " & p_CodEmp & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_NomEntFin = Trim(g_rst_Listas!EMPSUP_NOMCOR & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_TipoCreditoCtb(ByVal p_CodTip As String) As String
   moddat_gf_Consulta_TipoCreditoCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCCT WHERE "
   g_str_Parame = g_str_Parame & "TIPCCT_CODIGO = '" & p_CodTip & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_TipoCreditoCtb = Trim(g_rst_Listas!TIPCCT_DESCRI & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_SituacionCreditoCtb(ByVal p_ClaCre As String, p_CodSit As String) As String
   moddat_gf_Consulta_SituacionCreditoCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_SITCRE WHERE "
   g_str_Parame = g_str_Parame & "SITCRE_CLACRE = '" & p_ClaCre & "' AND "
   g_str_Parame = g_str_Parame & "SITCRE_CODSIT = '" & p_CodSit & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_SituacionCreditoCtb = Trim(g_rst_Listas!SITCRE_DESCRI & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_TipoCreditoCtb(p_Arregl() As moddat_tpo_Genera, p_Combo As ComboBox)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPCCT "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCCT_DESCRI ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         p_Combo.AddItem Trim(g_rst_Listas!TIPCCT_DESCRI)
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!TIPCCT_CODIGO)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!TIPCCT_DESCRI)
      
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_SituacionCreditoCtb(p_Arregl() As moddat_tpo_Genera, p_Combo As ComboBox, ByVal p_ClaCre As String)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_SITCRE WHERE "
   g_str_Parame = g_str_Parame & "SITCRE_CLACRE = '" & p_ClaCre & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SITCRE_DESCRI ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         p_Combo.AddItem Trim(g_rst_Listas!SITCRE_DESCRI)
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!SITCRE_CODSIT)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!SITCRE_DESCRI)
      
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_ConceptoCtb(ByVal p_CodCam As String) As String
   moddat_gf_Consulta_ConceptoCtb = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CONCTB WHERE "
   g_str_Parame = g_str_Parame & "CONCTB_CODCAM = '" & p_CodCam & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_ConceptoCtb = Trim(g_rst_Listas!CONCTB_DESCRI & "")
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Sub moddat_gs_Carga_ConceptoCtb(p_Arregl() As moddat_tpo_Genera, p_Combo As ComboBox, ByVal p_TipCon As Integer)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CONCTB WHERE "
   g_str_Parame = g_str_Parame & "CONCTB_TIPCON = " & CStr(p_TipCon) & " "
   g_str_Parame = g_str_Parame & "ORDER BY CONCTB_DESCRI ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         p_Combo.AddItem Trim(g_rst_Listas!CONCTB_DESCRI)
      
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!CONCTB_CODCAM)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!CONCTB_DESCRI)
      
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub moddat_gs_Carga_FerEve(p_Combo As ComboBox)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM COM_FEREVE WHERE "
   g_str_Parame = g_str_Parame & "FEREVE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY FEREVE_CODIGO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
         p_Combo.AddItem Trim(g_rst_Listas!FEREVE_NOMBRE)
         p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Listas!FEREVE_CODIGO)
         
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_FerEve(ByVal p_CodFer As Integer) As String
   moddat_gf_Consulta_FerEve = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM COM_FEREVE WHERE "
   g_str_Parame = g_str_Parame & "FEREVE_CODIGO = " & CStr(p_CodFer)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      moddat_gf_Consulta_FerEve = Trim(g_rst_Listas!FEREVE_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function moddat_gf_Consulta_ExposicionRCC(ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_TipMon As Integer, ByVal p_TipCam As Double, ByVal p_CuoFij As Double, ByVal p_IngTot As Double) As String
Dim r_int_ConTem        As Integer
Dim r_dbl_RccMna        As Double
Dim r_dbl_RccMex        As Double
Dim r_dbl_IngTot        As Double
Dim r_dbl_CuoFij        As Double
Dim r_dbl_ToDeAj        As Double
Dim r_dbl_IngNet        As Double
Dim r_dbl_CuMaMn        As Double
Dim r_dbl_CuMaMe        As Double
   
   moddat_gf_Consulta_ExposicionRCC = "NO"
   If p_TipMon = 2 Then
      r_int_ConTem = 6
      r_dbl_RccMna = 0: r_dbl_RccMex = 0: r_dbl_IngTot = 0: r_dbl_CuoFij = 0:
      r_dbl_ToDeAj = 0: r_dbl_IngNet = 0: r_dbl_CuMaMn = 0: r_dbl_CuMaMe = 0:
      
      If Not IsNull(p_CuoFij) Then
         r_dbl_CuoFij = p_CuoFij
      End If
      If Not IsNull(p_IngTot) Then
         r_dbl_IngTot = p_IngTot
      End If
      
      Call moddat_gs_Consulta_DeudaRCC(Trim(p_TipDoc), Trim(p_NumDoc), r_dbl_RccMna, r_dbl_RccMex)
      
      'TOTAL DEUDA AJUSTADA
      r_dbl_ToDeAj = (r_dbl_RccMna + (r_dbl_RccMex * IIf(r_int_ConTem < 4, 1.1, 1.2))) / IIf(r_int_ConTem < 4, 12 * r_int_ConTem, 12 * (r_int_ConTem - 3))
      
      'INGRESO NETO
      r_dbl_IngNet = r_dbl_IngTot - r_dbl_ToDeAj
      
      'CUOTA MAXIMA SOLES
      r_dbl_CuMaMn = r_dbl_IngNet * 0.4
      
      'CUOTA MAXIMA DOLARES
      r_dbl_CuMaMe = r_dbl_CuMaMn / (IIf(r_int_ConTem < 4, 1.1, 1.2) * p_TipCam)
      
      'EXPOSICION
      If r_dbl_CuMaMe < r_dbl_CuoFij Then
         moddat_gf_Consulta_ExposicionRCC = "SI"
      Else
         moddat_gf_Consulta_ExposicionRCC = "NO"
      End If
   End If
End Function

Public Function moddat_gf_Consulta_ExposicionGlobal(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_ValGar As Double, ByVal p_ValTGar As Integer, ByVal p_ValGti As Double) As Boolean
Dim r_dbl_ValCar     As Double
Dim r_dbl_ValRea     As Double
Dim r_dbl_ValHip     As Double
Dim r_dbl_ValGLi     As Double
Dim r_dbl_PatEfe     As Double

   moddat_gf_Consulta_ExposicionGlobal = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT (( "
   g_str_Parame = g_str_Parame & "            SELECT NVL(SUM(MAECFI_GARFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD <> '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                                            ) + "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(MAECFI_IMPFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 )"
   g_str_Parame = g_str_Parame & "          ) AS MONTO_GARANTIZADO , "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(EVATAS_VALREA_INM + EVATAS_VALREA_ES1 + EVATAS_VALREA_ES2 + EVATAS_VALREA_DE1 + EVATAS_VALREA_DE2), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "             WHERE EVATAS_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND EVATAS_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "          ) AS VALOR_REALIZACION   ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM( NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "             WHERE MAEGAR_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAEGAR_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAEGAR_TIPGAR = 1 "
   g_str_Parame = g_str_Parame & "               AND MAEGAR_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "          ) AS VALOR_GARANTIA_LIQUIDA    ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM( NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "             WHERE MAEGAR_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAEGAR_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAEGAR_TIPGAR = 2 "
   g_str_Parame = g_str_Parame & "               AND MAEGAR_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "          ) AS VALOR_GARANTIA_HIPOTECARIA    ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(CONLIM_PATEFE, 0) "
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2)"
   g_str_Parame = g_str_Parame & "          ) AS PATRIMONIO_EFECTIVO "
   g_str_Parame = g_str_Parame & "    FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      r_dbl_ValCar = CDbl(Trim$(g_rst_Genera!MONTO_GARANTIZADO)) + CDbl(p_ValGar)
      
      If p_ValTGar = 1 Then
         r_dbl_ValGLi = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_LIQUIDA)) + CDbl(p_ValGti)
         r_dbl_ValHip = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_HIPOTECARIA))
      ElseIf p_ValTGar = 2 Then
         r_dbl_ValHip = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_HIPOTECARIA)) + CDbl(p_ValGti)
         r_dbl_ValGLi = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_LIQUIDA))
      Else
         r_dbl_ValHip = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_HIPOTECARIA))
         r_dbl_ValGLi = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIA_LIQUIDA))
      End If
      
      r_dbl_ValRea = CDbl(Trim$(g_rst_Genera!VALOR_REALIZACION))
      r_dbl_PatEfe = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO))
      
      moddat_g_dbl_ValNv1 = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.1, 2)
      moddat_g_dbl_ValNv2 = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.15, 2)
      moddat_g_dbl_ValNv3 = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.3, 2)
      
      If Round(CDbl(r_dbl_ValCar), 2) <= CDbl(moddat_g_dbl_ValNv1) Then
         moddat_gf_Consulta_ExposicionGlobal = False
         
      ElseIf Round(CDbl(r_dbl_ValCar), 2) > CDbl(moddat_g_dbl_ValNv1) And Round(CDbl(r_dbl_ValCar), 2) <= CDbl(moddat_g_dbl_ValNv2) Then
         
         If r_dbl_ValHip > 0 And r_dbl_ValGLi = 0 Then                                                                     'únicamente con Garantía hipoteca
            If CDbl(r_dbl_ValRea) >= CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)) Then
               moddat_gf_Consulta_ExposicionGlobal = False
            Else
               moddat_gf_Consulta_ExposicionGlobal = True
               moddat_g_str_DesObs = Format(CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)), "###,###,##0.00")
            End If
            
         ElseIf r_dbl_ValHip = 0 And r_dbl_ValGLi > 0 Then                                                                 'únicamente con Garantía Líquida
            If r_dbl_ValGLi >= CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)) Then
               moddat_gf_Consulta_ExposicionGlobal = False
            Else
               moddat_gf_Consulta_ExposicionGlobal = True
               moddat_g_str_DesObs = Format(CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)), "###,###,##0.00")
            End If
            
         ElseIf r_dbl_ValHip > 0 And r_dbl_ValGLi > 0 Then                                                                 'garantía mixta, con Líquida e Hipotecaria
            If CDbl(CDbl(r_dbl_ValGLi) + CDbl(r_dbl_ValRea)) >= CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)) Then
               moddat_gf_Consulta_ExposicionGlobal = False
            Else
               moddat_gf_Consulta_ExposicionGlobal = True
               moddat_g_str_DesObs = Format(CDbl(CDbl(r_dbl_ValCar) - CDbl(moddat_g_dbl_ValNv1)), "###,###,##0.00")
            End If
            
         Else                                                                                                              'No tiene garantía líquida ni hipotecaria
            moddat_gf_Consulta_ExposicionGlobal = True
         End If
         
      ElseIf Round(CDbl(r_dbl_ValCar), 2) > CDbl(moddat_g_dbl_ValNv2) And Round(CDbl(r_dbl_ValCar), 2) <= CDbl(moddat_g_dbl_ValNv3) Then
         'Con Hipoteca Mixta
         If r_dbl_ValHip > 0 And r_dbl_ValGLi > 0 Then
            If Round(CDbl(r_dbl_ValRea), 2) >= CDbl(0.05 * CDbl(r_dbl_PatEfe)) Then                                        'garantía mixta con hipoteca mayor e igual al 5%.PE
               If r_dbl_ValGLi >= (Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(0.05 * CDbl(r_dbl_PatEfe)), 2)) - Round(CDbl(r_dbl_ValRea), 2) Then
                  moddat_gf_Consulta_ExposicionGlobal = False
               Else
                  moddat_gf_Consulta_ExposicionGlobal = True
                  moddat_g_str_DesObs = Format((Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(0.05 * CDbl(r_dbl_PatEfe)), 2)) - Round(CDbl(r_dbl_ValRea), 2), "###,###,##0.00")
               End If
            ElseIf Round(CDbl(r_dbl_ValRea), 2) < CDbl(0.05 * CDbl(r_dbl_PatEfe)) Then                                     'garantía mixta con hipoteca menor al 5%.PE
               If Round(CDbl(r_dbl_ValGLi), 2) >= Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(r_dbl_ValRea), 2) - Round(CDbl(moddat_g_dbl_ValNv1), 2) Then
                  moddat_gf_Consulta_ExposicionGlobal = False
               Else
                  moddat_gf_Consulta_ExposicionGlobal = True
                  moddat_g_str_DesObs = Format(Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(r_dbl_ValRea), 2) - Round(CDbl(moddat_g_dbl_ValNv1), 2), "###,###,##0.00")
               End If
            End If
            
         'Con garantía única
         Else
            If Round(CDbl(r_dbl_ValGLi), 2) >= Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(moddat_g_dbl_ValNv1), 2) Then
               moddat_gf_Consulta_ExposicionGlobal = False
            Else
               moddat_gf_Consulta_ExposicionGlobal = True
               moddat_g_str_DesObs = Format(Round(CDbl(r_dbl_ValCar), 2) - Round(CDbl(moddat_g_dbl_ValNv1), 2), "###,###,##0.00")
            End If
         End If
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function

Public Function moddat_gf_Consulta_ExposicionGlobal_OLD(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_ValGar As Double, ByVal p_ValGti) As Boolean
'ByVal p_PerMes As Integer, ByVal p_PerAno As Integer,
Dim r_dbl_ValCar    As Double
Dim r_dbl_ValExp    As Double
Dim r_dbl_ValGar    As Double

   moddat_gf_Consulta_ExposicionGlobal_OLD = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT (( "
   g_str_Parame = g_str_Parame & "            SELECT NVL(SUM(MAECFI_GARFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD <> '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                                            ) + "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(MAECFI_IMPFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 )"
   g_str_Parame = g_str_Parame & "          ) AS MONTO_GARANTIZADO , "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(EVATAS_VALREA_INM + EVATAS_VALREA_ES1 + EVATAS_VALREA_ES2 + EVATAS_VALREA_DE1 + EVATAS_VALREA_DE2), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "             WHERE EVATAS_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND EVATAS_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "          ) AS VALOR_REALIZACION   ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM( NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "             WHERE MAEGAR_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "               AND MAEGAR_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAEGAR_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "          ) AS VALOR_GARANTIAS    ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(CONLIM_PATEFE, 0) "
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2)"
   g_str_Parame = g_str_Parame & "          ) AS PATRIMONIO_EFECTIVO "
   g_str_Parame = g_str_Parame & "    FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_dbl_ValCar = CDbl(Trim$(g_rst_Genera!MONTO_GARANTIZADO)) + CDbl(p_ValGar)
      r_dbl_ValExp = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.1
      r_dbl_ValGar = CDbl(Trim$(g_rst_Genera!VALOR_GARANTIAS)) + CDbl(p_ValGti)
      
      If Round(CDbl(CDbl(r_dbl_ValExp) + CDbl(r_dbl_ValGar)), 2) < CDbl(r_dbl_ValCar) Then
         moddat_gf_Consulta_ExposicionGlobal_OLD = True
      End If
      
      moddat_g_dbl_ValSGa = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.1, 2)
      moddat_g_dbl_ValGHi = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.15, 2)
      moddat_g_dbl_ValGLi = Round(CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.3, 2)
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
End Function

Public Function moddat_gf_Consulta_NivelEndeudamiento(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_ValGar As Double) As Boolean
Dim r_dbl_ValCGar    As Double
Dim r_dbl_ValExpo    As Double

'   'Año y mes
'   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)

   moddat_gf_Consulta_NivelEndeudamiento = False
         
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT (( "
   g_str_Parame = g_str_Parame & "            SELECT NVL(SUM(MAECFI_GARFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & IIf(p_TipDoc = 7, 6, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD <> '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                                            ) + "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(MAECFI_IMPFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & IIf(p_TipDoc = 6, 7, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 )"
   g_str_Parame = g_str_Parame & "          ) AS MONTO_GARANTIZADO , "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(EVATAS_VALREA_INM + EVATAS_VALREA_ES1 + EVATAS_VALREA_ES2 + EVATAS_VALREA_DE1 + EVATAS_VALREA_DE2), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "             WHERE EVATAS_TIPDOC = " & IIf(p_TipDoc = 7, 6, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND EVATAS_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "          ) AS VALOR_REALIZACION   ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(CONLIM_PATEFE, 0) "
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2)"
   g_str_Parame = g_str_Parame & "          ) AS PATRIMONIO_EFECTIVO "
   g_str_Parame = g_str_Parame & "    FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_dbl_ValCGar = CDbl(Trim$(g_rst_Genera!MONTO_GARANTIZADO)) + CDbl(p_ValGar) '- CDbl(Trim$(g_rst_Genera!VALOR_REALIZACION))
      r_dbl_ValExpo = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.3
      
      If CDbl(r_dbl_ValCGar) - CDbl(r_dbl_ValExpo) > 0 Then
         moddat_gf_Consulta_NivelEndeudamiento = True
      Else
         moddat_gf_Consulta_NivelEndeudamiento = False
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
End Function

Public Function moddat_gf_Consulta_NivelEndeudamiento_Old(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_ValGar As Double) As Boolean
Dim r_dbl_ValCGar    As Double
Dim r_dbl_ValExpo    As Double

'   'Año y mes
'   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)

   moddat_gf_Consulta_NivelEndeudamiento_Old = False
         
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT (( "
   g_str_Parame = g_str_Parame & "            SELECT NVL(SUM(MAECFI_GARFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & IIf(p_TipDoc = 7, 6, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                                                                              ) + "
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(COMCIE_TOTPRE), 0) "
   g_str_Parame = g_str_Parame & "              FROM CRE_COMCIE "
   g_str_Parame = g_str_Parame & "             WHERE COMCIE_PERMES = " & p_PerMes & ""
   g_str_Parame = g_str_Parame & "               AND COMCIE_PERANO = " & p_PerAno & ""
   g_str_Parame = g_str_Parame & "               AND COMCIE_TDOCLI = " & IIf(p_TipDoc = 6, 7, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND COMCIE_NDOCLI = '" & p_NumDoc & "' )) AS MONTO_GARANTIZADO , "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(EVATAS_VALREA_INM + EVATAS_VALREA_ES1 + EVATAS_VALREA_ES2 + EVATAS_VALREA_DE1 + EVATAS_VALREA_DE2), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "             WHERE EVATAS_TIPDOC = " & IIf(p_TipDoc = 7, 6, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND EVATAS_NUMDOC = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "          )+ "
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(COMCIE_MTOREA), 0) "
   g_str_Parame = g_str_Parame & "              FROM CRE_COMCIE"
   g_str_Parame = g_str_Parame & "             WHERE COMCIE_PERMES = " & p_PerMes & ""
   g_str_Parame = g_str_Parame & "               AND COMCIE_PERANO = " & p_PerAno & ""
   g_str_Parame = g_str_Parame & "               AND COMCIE_TDOCLI = " & IIf(p_TipDoc = 6, 7, p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "               AND COMCIE_NDOCLI = '" & p_NumDoc & "'"
   g_str_Parame = g_str_Parame & "          ) AS VALOR_REALIZACION   ,"
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(CONLIM_PATEFE, 0) "
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2)"
   g_str_Parame = g_str_Parame & "          ) AS PATRIMONIO_EFECTIVO "
   g_str_Parame = g_str_Parame & "    FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      r_dbl_ValCGar = CDbl(Trim$(g_rst_Genera!MONTO_GARANTIZADO)) + CDbl(p_ValGar) '- CDbl(Trim$(g_rst_Genera!VALOR_REALIZACION))
      r_dbl_ValExpo = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.3
      
      If CDbl(r_dbl_ValCGar) - CDbl(r_dbl_ValExpo) > 0 Then
         moddat_gf_Consulta_NivelEndeudamiento_Old = True
      Else
         moddat_gf_Consulta_NivelEndeudamiento_Old = False
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
End Function

Public Sub moddat_gs_Consulta_DeudaRCC(p_TipDoc As String, p_NumDoc As String, ByRef p_RccMna As Double, ByRef p_RccMex As Double)
Dim r_str_Cadena   As String

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT * "
   r_str_Cadena = r_str_Cadena & "  FROM CLI_RCCDET "
   r_str_Cadena = r_str_Cadena & " WHERE RCCDET_TIPDOC = '" & p_TipDoc & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_NUMDOC = '" & p_NumDoc & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_PERMES = (SELECT RCCCAB_PERMES FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES"
   r_str_Cadena = r_str_Cadena & "                                                     FROM CLI_RCCCAB ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) WHERE ROWNUM < 2)"
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_PERANO = (SELECT RCCCAB_PERANO FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES"
   r_str_Cadena = r_str_Cadena & "                                                     FROM CLI_RCCCAB ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) WHERE ROWNUM < 2)"
   
   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      If g_rst_Listas!RCCDET_MONDEU = 1 Then
         p_RccMna = p_RccMna + Format(g_rst_Listas!RCCDET_MTOSOL, "###,###,##0.00")
      ElseIf g_rst_Listas!RCCDET_MONDEU = 2 Then
         p_RccMex = p_RccMex + Format(g_rst_Listas!RCCDET_MTODOL, "###,###,##0.00")
      End If
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Function moddat_gf_Consulta_SobreEndeudamiento(ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_PerMes As String, ByVal p_PerAno As String) As String
Dim r_str_Cadena   As String

   moddat_gf_Consulta_SobreEndeudamiento = "0"
   
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT RCCCAB_DEUCA0, RCCCAB_DEUCA1, RCCCAB_DEUCA2, RCCCAB_DEUCA3, RCCCAB_DEUCA4 "
   r_str_Cadena = r_str_Cadena & "  FROM CLI_RCCCAB "
   r_str_Cadena = r_str_Cadena & " WHERE RCCCAB_TIPDOC = " & p_TipDoc & " "
   r_str_Cadena = r_str_Cadena & "   AND RCCCAB_NUMDOC = '" & Trim(p_NumDoc) & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCCAB_PERMES = " & p_PerMes & " "
   r_str_Cadena = r_str_Cadena & "   AND RCCCAB_PERANO = " & p_PerAno & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      If g_rst_Listas!RCCCAB_DEUCA2 + g_rst_Listas!RCCCAB_DEUCA3 + g_rst_Listas!RCCCAB_DEUCA4 > 1000 Then
         moddat_gf_Consulta_SobreEndeudamiento = "1"
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function


' 27012020 RAT INICIO
Public Sub fs_Envia_CorreoInspektor(p_mpsSesion As MAPISession, p_mpsMsj As MAPIMessages, p_Asunto As String, p_Mensaje As String, Optional ByVal p_Correos As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
Dim r_str_Cadena     As String
Dim r_str_CadArr()   As String
Dim r_int_Contar     As Integer

   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
'   If (Len(Trim(p_User1)) > 0) Then
'       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
'       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
'   End If
'
'   'Ejecutivo de Seguimientos
'   If (Len(Trim(p_User2)) > 0) Then
'       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
'       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
'          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
'          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
'       End If
'   End If
'
   'If (Len(Trim(p_NumSol)) > 0) Then
   '    r_str_Cadena = moddat_gf_UsuObs(p_NumSol, p_CodIns)
   '    ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   '    moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_UsuSis(r_str_Cadena)
   'End If
   
   'Director de Producción
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)
  moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(900, moddat_g_arr_Genera)
    moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(902, moddat_g_arr_Genera)
   'Director Comercial
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(100, moddat_g_arr_Genera)
   
   'Jefe de Ventas
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(120, moddat_g_arr_Genera)
   
   'Jefe de Seguimiento
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(130, moddat_g_arr_Genera)
   
   'Jefe de Operaciones
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)
   
   'Evaluador de Operaciones
'   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
   
   'Legal
'   If (p_Legal = True) Then
'      'Jefe de Legal
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
'
'      'Asistente de Legal 1
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)
'
'      'Asistente de Legal 2
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)
'   End If
   
   'Creditos
'   If (p_JfCred = True) Then
'      'Jefe de Créditos
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(210, moddat_g_arr_Genera)
'
'      'Evaluadores de credito
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_UsuEje_Arr(211, moddat_g_arr_Genera, moddat_g_str_NumSol)
'   End If
   
'   If (p_DrAdm = True) Then
'      'Director de Administración
'      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(300, moddat_g_arr_Genera)
'   End If
'
   'Director General
   'moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(10, moddat_g_arr_Genera)
   
   'Separa correos por ";"
   If Trim(p_Correos & "") <> "" Then
      r_str_CadArr = Split(p_Correos, ";")
      For r_int_Contar = LBound(r_str_CadArr) To UBound(r_str_CadArr)
          If Trim(r_str_CadArr(r_int_Contar)) <> "" Then
             ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
             moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_CadArr(r_int_Contar)
          End If
      Next
   End If
   Call moddat_gs_EnvCor(p_mpsSesion, p_mpsMsj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, p_NomFil, p_RutFil)
End Sub
' 27012020 RAT FIN



Public Sub fs_Envia_CorreoEle(p_mpsSesion As MAPISession, p_mpsMsj As MAPIMessages, p_Asunto As String, p_Mensaje As String, p_User1 As String, p_User2 As String, p_NumSol As String, p_CodIns As Integer, p_JfCred As Boolean, p_Legal As Boolean, p_DrAdm As Boolean, Optional ByVal p_Correos As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
Dim r_str_Cadena     As String
Dim r_str_CadArr()   As String
Dim r_int_Contar     As Integer

   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   If (Len(Trim(p_User1)) > 0) Then
       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
   End If
   
   'Ejecutivo de Seguimientos
   If (Len(Trim(p_User2)) > 0) Then
       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
       End If
   End If
   
   'If (Len(Trim(p_NumSol)) > 0) Then
   '    r_str_Cadena = moddat_gf_UsuObs(p_NumSol, p_CodIns)
   '    ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   '    moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_UsuSis(r_str_Cadena)
   'End If
   
   'Director de Producción
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)
   
   'Director Comercial
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(100, moddat_g_arr_Genera)
   
   'Jefe de Ventas
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(120, moddat_g_arr_Genera)
   
   'Jefe de Seguimiento
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(130, moddat_g_arr_Genera)
   
   'Jefe de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)
   
   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
   
   'Legal
   If (p_Legal = True) Then
      'Jefe de Legal
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
      
      'Asistente de Legal 1
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)
      
      'Asistente de Legal 2
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)
   End If
   
   'Creditos
   If (p_JfCred = True) Then
      'Jefe de Créditos
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(210, moddat_g_arr_Genera)
   
      'Evaluadores de credito
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_UsuEje_Arr(211, moddat_g_arr_Genera, moddat_g_str_NumSol)
   End If
   
   If (p_DrAdm = True) Then
      'Director de Administración
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(300, moddat_g_arr_Genera)
   End If
   
   'Director General
   'moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(10, moddat_g_arr_Genera)
   
   'Separa correos por ";"
   If Trim(p_Correos & "") <> "" Then
      r_str_CadArr = Split(p_Correos, ";")
      For r_int_Contar = LBound(r_str_CadArr) To UBound(r_str_CadArr)
          If Trim(r_str_CadArr(r_int_Contar)) <> "" Then
             ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
             moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_CadArr(r_int_Contar)
          End If
      Next
   End If
   
   Call moddat_gs_EnvCor(p_mpsSesion, p_mpsMsj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, p_NomFil, p_RutFil)
End Sub

Public Sub fs_Envia_Correo_Prom(p_mpsSesion As MAPISession, p_mpsMsj As MAPIMessages, p_Asunto As String, p_Mensaje As String, p_User1 As String, p_User2 As String, p_JfCred As Boolean, p_Legal As Boolean, p_DrAdm As Boolean, p_TpOpc As Boolean, p_EvaOpe As Boolean, p_JfCtb As Boolean)
Dim r_str_Cadena     As String

   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   If (Len(Trim(p_User1)) > 0) Then
       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
   End If
   
   'Ejecutivo de Seguimientos
   If (Len(Trim(p_User2)) > 0) Then
       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
       End If
   End If
   
   'Director de Producción
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)
       
   If p_TpOpc = True Then
      'Director Comercial
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(100, moddat_g_arr_Genera)
      
      'Jefe de Ventas
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(120, moddat_g_arr_Genera)
      
      'Jefe de Seguimiento
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(130, moddat_g_arr_Genera)
   End If
   
   If p_EvaOpe = True Then
      'Jefe de Operaciones
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)
        
      'Evaluador de Operaciones
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
   End If
   
   If p_Legal = True Then
      'Jefe de Legal
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)
      
      'Asistente de Legal 1
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)
      
      'Asistente de Legal 2
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)
   End If
   
   If p_JfCred = True Then
      'Jefe de Créditos
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(210, moddat_g_arr_Genera)
   End If
   
   If p_DrAdm = True Then
      'Director de Administración
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(300, moddat_g_arr_Genera)
   End If
   
   If p_JfCtb = True Then
      'Jefe de Contabilidad
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(310, moddat_g_arr_Genera)
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(134, moddat_g_arr_Genera)
   End If
   
   'Director General
   'moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(10, moddat_g_arr_Genera)
   
   Call moddat_gs_EnvCor(p_mpsSesion, p_mpsMsj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub

Public Sub fs_Envia_CorreoOpe(p_mpsSesion As MAPISession, p_mpsMsj As MAPIMessages, p_Asunto As String, p_Mensaje As String, p_User1 As String, p_User2 As String, p_NumSol As String, p_CodIns As Integer, Optional ByVal p_Correos As String, Optional ByVal p_NomFil As String, Optional ByVal p_RutFil As String)
Dim r_str_Cadena     As String
Dim r_str_CadArr()   As String
Dim r_int_Contar     As Integer

   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   If (Len(Trim(p_User1)) > 0) Then
       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
   End If
   
   'Ejecutivo de Seguimientos
   If (Len(Trim(p_User2)) > 0) Then
       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
       End If
   End If
   
   'Director de Producción
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)
   
   'Jefe de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)
   
   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)
   
   
   'Separa correos por ";"
   If Trim(p_Correos & "") <> "" Then
      r_str_CadArr = Split(p_Correos, ";")
      For r_int_Contar = LBound(r_str_CadArr) To UBound(r_str_CadArr)
          If Trim(r_str_CadArr(r_int_Contar)) <> "" Then
             ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
             moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_CadArr(r_int_Contar)
          End If
      Next
   End If
   
   Call moddat_gs_EnvCor(p_mpsSesion, p_mpsMsj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, p_NomFil, p_RutFil)
End Sub

Public Function fs_EnviarCorreoAdj(p_Sesion As MAPISession, p_Mensaje As MAPIMessages, p_Asunto As String, p_Contenido As String, p_User1 As String, p_User2 As String, p_JfCred As Boolean, p_Legal As Boolean, p_DrAdm As Boolean) As String
Dim r_int_Contad      As Integer
Dim r_int_Index       As Integer
Dim r_str_Cadena      As String
   
   On Error GoTo moddat_gf_EnvCor
      
   fs_EnviarCorreoAdj = ""
   r_str_Cadena = ""
   
   'Inicializa
   p_Sesion.DownLoadMail = False
   p_Sesion.NewSession = True
   p_Sesion.SignOn
   p_Mensaje.SessionID = p_Sesion.SessionID
  
   'Envío
   p_Mensaje.Compose
   '-----------------------------------------------------------------------------------------------------
   
   'Consejero Hipotecario
   If (Len(Trim(p_User1)) > 0) Then
       ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
       moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User1))
   End If

   'Ejecutivo de Seguimientos
   If (Len(Trim(p_User2)) > 0) Then
       r_str_Cadena = moddat_gf_Buscar_DirEle_Codigo(Trim(p_User2))
       If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
          ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
          moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
       End If
   End If

   'Director de Producción
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(200, moddat_g_arr_Genera)

   'Director Comercial
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(100, moddat_g_arr_Genera)

   'Jefe de Ventas
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(120, moddat_g_arr_Genera)

   'Jefe de Seguimiento
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(130, moddat_g_arr_Genera)

   'Jefe de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(220, moddat_g_arr_Genera)

   'Evaluador de Operaciones
   moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(221, moddat_g_arr_Genera)

   'Legal
   If (p_Legal = True) Then
      'Jefe de Legal
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(230, moddat_g_arr_Genera)

      'Asistente de Legal 1
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(231, moddat_g_arr_Genera)

      'Asistente de Legal 2
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(232, moddat_g_arr_Genera)
   End If

   'Creditos
   If (p_JfCred = True) Then
      'Jefe de Créditos
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(210, moddat_g_arr_Genera)

      'Evaluadores de credito
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_UsuEje_Arr(211, moddat_g_arr_Genera, moddat_g_str_NumSol)
   End If

   If (p_DrAdm = True) Then
      'Director de Administración
      moddat_g_arr_Genera = moddat_gf_Buscar_DirEle_TipUsu_Arr(300, moddat_g_arr_Genera)
   End If
   '-----------------------------------------------------------------------------------------------------
   
   For r_int_Contad = 0 To UBound(moddat_g_arr_Genera) - 1
      If Len(Trim(moddat_g_arr_Genera(r_int_Contad + 1).Genera_Codigo)) > 0 Then
         p_Mensaje.RecipIndex = r_int_Contad
         p_Mensaje.RecipDisplayName = moddat_g_arr_Genera(r_int_Contad + 1).Genera_Codigo
      End If
   Next r_int_Contad
   
   p_Mensaje.MsgSubject = p_Asunto
   p_Mensaje.MsgNoteText = p_Contenido

   For r_int_Contad = 0 To UBound(moddat_g_arr_GenAux) - 1
      If Len(Trim(moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Codigo)) > 0 Then
         p_Mensaje.AttachmentIndex = r_int_Contad
         p_Mensaje.AttachmentName = moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Codigo  'p_NomFil_01
         p_Mensaje.AttachmentPathName = moddat_g_arr_GenAux(r_int_Contad + 1).Genera_Refere 'p_RutFil_01
         p_Mensaje.AttachmentPosition = r_int_Contad
         p_Mensaje.AttachmentType = mapData
      End If
   Next r_int_Contad
   
   'Enviar Correro
   p_Mensaje.Send
   DoEvents
   fs_EnviarCorreoAdj = ""
   
  'Cierra la sesión
  p_Sesion.SignOff
  Exit Function
  
moddat_gf_EnvCor:
   fs_EnviarCorreoAdj = Err.Description
   p_Sesion.SignOff
   MsgBox Err.Description, vbCritical
End Function

Public Function moddat_gf_Buscar_DirEle_TipUsu_Arr(ByVal p_TipUsu As Integer, ByRef p_Arreglo() As moddat_tpo_Genera) As moddat_tpo_Genera()
   g_str_Parame = "SELECT * FROM CRE_EJECMC A, CRE_EJETIP B WHERE "
   g_str_Parame = g_str_Parame & "A.EJECMC_CODEJE = B.EJETIP_CODEJE AND "
   g_str_Parame = g_str_Parame & "A.EJECMC_SITUAC <> 2 AND "
   g_str_Parame = g_str_Parame & "B.EJETIP_TIPEJE = " & CStr(p_TipUsu) & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      Do While Not g_rst_Genera.EOF
         If Not moddat_gf_Verifica_DirEle(p_Arreglo, Trim(g_rst_Genera!EJECMC_DIRELE & "")) Then
            ReDim Preserve p_Arreglo(UBound(p_Arreglo) + 1)
            p_Arreglo(UBound(p_Arreglo)).Genera_Codigo = Trim(g_rst_Genera!EJECMC_DIRELE & "")
            p_Arreglo(UBound(p_Arreglo)).Genera_Nombre = Trim(g_rst_Genera!EJECMC_CODEJE & "")
         End If
         
         g_rst_Genera.MoveNext
      Loop
   End If

   moddat_gf_Buscar_DirEle_TipUsu_Arr = p_Arreglo
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Public Function moddat_gf_Buscar_DirEle_UsuEje_Arr(ByVal p_TipUsu As Integer, ByRef p_Arreglo() As moddat_tpo_Genera, p_NumSol As String) As moddat_tpo_Genera()
Dim r_arr_Arreglo()  As moddat_tpo_Genera
Dim r_int_Contar     As Integer
Dim r_rst_PriAux     As ADODB.Recordset
Dim r_str_Parame     As String
   
   If Trim(p_NumSol) <> "" Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT SEGUSUCRE  "
      r_str_Parame = r_str_Parame & "   FROM TRA_SEGDET A  "
      r_str_Parame = r_str_Parame & "  WHERE A.SEGDET_NUMSOL = '" & p_NumSol & "' "
      r_str_Parame = r_str_Parame & "    AND A.SEGDET_CODOCU IN (11,15,17,18,21)  "

      If Not gf_EjecutaSQL(r_str_Parame, r_rst_PriAux, 3) Then
         Exit Function
      End If
   
      If Not (r_rst_PriAux.BOF And r_rst_PriAux.EOF) Then
         ReDim r_arr_Arreglo(0)
         r_arr_Arreglo = moddat_gf_Buscar_DirEle_TipUsu_Arr(211, r_arr_Arreglo)
         For r_int_Contar = 0 To UBound(r_arr_Arreglo)
             If Trim(r_rst_PriAux!SEGUSUCRE & "") = Trim(r_arr_Arreglo(r_int_Contar).Genera_Nombre) Then
                ReDim Preserve p_Arreglo(UBound(p_Arreglo) + 1)
                p_Arreglo(UBound(p_Arreglo)).Genera_Codigo = Trim(r_arr_Arreglo(r_int_Contar).Genera_Codigo)
                Exit For
             End If
         Next
      End If
      r_rst_PriAux.Close
      Set r_rst_PriAux = Nothing
      
      moddat_gf_Buscar_DirEle_UsuEje_Arr = p_Arreglo
   End If
End Function

Public Sub moddat_gf_Cargar_AgrPrd()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM CRE_AGRPRO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
            
      Do While Not g_rst_Genera.EOF
         Select Case Trim(g_rst_Genera!AGRPRO_CODAGR)
            Case "AGRCRC": moddat_g_str_AgrCRC = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGRCME": moddat_g_str_AgrCME = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGRTMIC": moddat_g_str_AgrTMIC = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGR1MIC": moddat_g_str_Agr1MIC = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGR2MIC": moddat_g_str_Agr2MIC = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGRTFMV": moddat_g_str_AgrTFMV = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGRMIHG": moddat_g_str_AgrMIHG = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGR1FMV": moddat_g_str_Agr1FMV = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGR2FMV": moddat_g_str_Agr2FMV = Trim(g_rst_Genera!AGRPRO_DESCRI)
            Case "AGRCOF": moddat_g_str_AgrCOF = Trim(g_rst_Genera!AGRPRO_DESCRI)
         End Select
         g_rst_Genera.MoveNext
         
         If g_rst_Genera.EOF Then
            Exit Do
         End If
      Loop
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub moddat_gs_Consulta_DatBMS(ByVal p_CodPry As String, ByRef p_FlgAfeBV As Integer, ByRef p_TipAfeBV As Integer, ByRef p_ValAfeBV As Double)
   p_FlgAfeBV = 0
   p_TipAfeBV = 0
   p_ValAfeBV = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DATGEN_FLGAFEBV, DATGEN_FLGTIPAFE, DATGEN_VALAFEBV "
   g_str_Parame = g_str_Parame & "   FROM PRY_DATGEN  "
   g_str_Parame = g_str_Parame & "  WHERE DATGEN_SITUAC = 1 AND DATGEN_CODIGO = '" & p_CodPry & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   p_FlgAfeBV = IIf(IsNull(g_rst_Genera!DATGEN_FLGAFEBV), 0, Trim(g_rst_Genera!DATGEN_FLGAFEBV & ""))
   p_TipAfeBV = IIf(IsNull(g_rst_Genera!DATGEN_FLGTIPAFE), 0, Trim(g_rst_Genera!DATGEN_FLGTIPAFE & ""))
   p_ValAfeBV = IIf(IsNull(g_rst_Genera!DATGEN_VALAFEBV), 0, Trim(g_rst_Genera!DATGEN_VALAFEBV & ""))
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Function moddat_gf_Consulta_ParTasInt(p_Arregl() As moddat_tpo_Genera, ByVal p_CodPrd As String, ByVal p_CodPry As String, ByVal p_ValPre_Min As Double, ByVal p_ValPre_Max As Double, ByVal p_ValInm_Min As Double, _
                                             ByVal p_ValInm_Max As Double, ByVal p_PlzPre_Min As Integer, ByVal p_PlzPre_Max As Integer, ByVal p_PorIni_Min As Double, ByVal p_PorIni_Max As Double, ByVal p_TipBon As Integer) As Integer
Dim p_TipEva As Integer
   
   p_TipEva = 0
   moddat_gf_Consulta_ParTasInt = False
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_TASINT ("
   g_str_Parame = g_str_Parame & Month(Format(Now, "dd/mm/yyyy")) & ", "
   g_str_Parame = g_str_Parame & Year(Format(Now, "dd/mm/yyyy")) & " , "
   g_str_Parame = g_str_Parame & "'REPORTE TASA DE INTERES' , "
   g_str_Parame = g_str_Parame & "'" & CStr(Format(p_CodPrd, "000")) & "' , "
   g_str_Parame = g_str_Parame & p_TipEva & ", "
   g_str_Parame = g_str_Parame & "'" & CStr(Format(p_CodPry, "000000")) & "' , "
   g_str_Parame = g_str_Parame & p_TipBon & ", "
   g_str_Parame = g_str_Parame & 1 & ", "
   g_str_Parame = g_str_Parame & CDbl(p_ValPre_Min) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_ValPre_Max) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_ValInm_Min) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_ValInm_Max) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_PlzPre_Min) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_PlzPre_Max) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_PorIni_Min) & ", "
   g_str_Parame = g_str_Parame & CDbl(p_PorIni_Max) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "') "
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_gf_Consulta_ParTasInt = True
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
         p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim$(g_rst_Princi!CODIGO)
         p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim$(g_rst_Princi!DESCRI)
         p_Arregl(UBound(p_Arregl)).Genera_Cantid = g_rst_Princi!TASINT
         g_rst_Princi.MoveNext
      Loop
      
      If UBound(p_Arregl) > 1 Then
          MsgBox "Se encontró más de un registro.", vbExclamation, modgen_g_str_NomPlt
      End If
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Function

