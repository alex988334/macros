Const mapName = "zulu://10.0.2.20:6473/ТЕСТ/Test_Gorelik/Алексей_map.zmp"

Dim VO_SU, VO_NU, VO_K, VO_PR, VS_IV, VS_UVS, VS_U, VS_VK, VS_POT, VS_PR, GS_U, GS_UCH, GS_POT, GS_PR, DK_K, DK_SU, DK_NU, DK_PR
Dim SS_U, SS_UCH, SS_POT, SS_PR, TS_I, TS_UCH, TS_U, TS_OP, TS_PR, ES_POD, ES_RPTPPP, ES_UCH, ES_U, ES_AB, ES_PR 
Dim  NetKeys, Nets, VO, VS, GS, DK, SS, TS, ES   

    init()
    VO_SU = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","D_fixed_с","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VO_NU = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","D","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VO_K = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","N_T_P","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VO_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")   
    
    VS_IV = Array("sys","G_KSIO_name","G_Etap","G_Meropriyatiya","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VS_UVS = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","D","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VS_U = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","G_point_in","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VS_VK = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","G_point_in","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VS_POT = Array("sys","G_KSIO_name","V_Etap","V_Nov_Rek_Dem","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    VS_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")
    
    GS_U = Array("sys","G_KSIO_name","K_Etap","K_Meropriyatiya","G_point_in","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    GS_UCH = Array("sys","G_KSIO_name","K_Etap","K_Meropriyatiya","diam","V_L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    GS_POT = Array("sys","G_KSIO_name","K_Etap","K_Meropriyatiya","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    GS_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")	     
  
	DK_K = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    DK_SU = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","D_fixed_с","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    DK_NU = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","D","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    DK_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")
    
    SS_U = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","G_point_in","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	
    SS_UCH = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","L","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	
    SS_POT = Array("sys","G_KSIO_name","N_Etap","N_Meropriyatiya","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	
    SS_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")	
    
    TS_I = Array("sys","G_KSIO_name","G_Etap","G_Meropriyatiya","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    TS_UCH = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","L","Dpod","Dobr","V_Vid seti","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    TS_U = Array("sys","G_KSIO_name","V_Etap","V_Meropriyatiya","G_point_in","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    TS_OP = Array("sys","G_KSIO_name","V_Etap","V_Nov_Rek_Dem","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")
    TS_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")
    
    ES_POD = Array("sys","G_KSIO_name","Этап строительства","Назначение","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    ES_RPTPPP = Array("sys","G_KSIO_name","Этап строительства","Назначение","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    ES_UCH = Array("sys","G_KSIO_name","Этап","Назначение кабеля","Класс напряжения","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    ES_U = Array("sys","G_KSIO_name","Этап","Назначение","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    ES_AB = Array("sys","G_KSIO_name","Этап","Назначение","G_Num_obrash","G_obj_po_obrash","G_layer_etap","G_layer_sito")	 
    ES_PR = Array("sys","G_KSIO_name","G_Num_obrash","G_layer_etap","G_layer_sito")	 

Const NET_VO = "ВО"
Const NET_VS = "ВС"
Const NET_GS = "ГС"
Const NET_DK = "ДК"
Const NET_SS = "СС"
Const NET_TS = "ТС"
Const NET_ES = "ЭС"
Const DELIMETER = "_"
NetKeys = Array(NET_VO, NET_VS, NET_GS, NET_DK, NET_SS, NET_TS, NET_ES)
		
Const SAM_UCH = "Самотёчный участок" 
Const NAP_UCH = "Напорный участок"
Const KOL = "Колодец"
Const PRIM = "Примитив"  
Const IST_VOD = "Источник водоснабжения"
Const UCH_VOD_SET = "Участок водопроводной сети"
Const UZEL = "Узел"	
Const VOD_KOL_GIDRAN = "Водопроводный колодец с гидрантом"
Const POT = "Потребитель"
Const UCH = "Участок"
Const IST = "Источник"
Const OBOB_POT = "Обобщенный потребитель"
Const PODST = "Подстанция"
Const RP_TP_PP = "РП_ТП_ПП"
Const UCHASTKI = "Участки"
Const ABON = "Абонент"

Set VO = new Map
Call VO.init(Array("Самотёчный участок", "Напорный участок", "Колодец", "Примитив"), Array(VO_SU, VO_NU, VO_K, VO_PR))
Set VS = new Map
Call VS.init(Array("Источник водоснабжения", "Участок водопроводной сети", "Узел", "Водопроводный колодец с гидрантом", "Потребитель", "Примитив"), Array(VS_IV, VS_UVS, VS_U, VS_VK, VS_POT, VS_PR))
Set GS = new Map
Call GS.init(Array("Узел", "Участок", "Потребитель", "Примитив"), Array(GS_U, GS_UCH, GS_POT, GS_PR))
Set DK = new Map
Call DK.init(Array("Колодец", "Самотёчный участок", "Напорный участок", "Примитив"), Array(DK_K, DK_SU, DK_NU, DK_PR))
Set SS = new Map
Call SS.init(Array("Узел", "Участок", "Потребитель", "Примитив"), Array(SS_U, SS_UCH, SS_POT, SS_PR))
Set TS = new Map
Call TS.init(Array("Источник", "Участок", "Узел", "Обобщенный потребитель", "Примитив"), Array(TS_I, TS_UCH, TS_U, TS_OP, TS_PR))
Set ES = new Map 
Call ES.init(Array("Подстанция", "РП_ТП_ПП", "Участки", "Узел", "Абонент", "Примитив"), Array(ES_POD, ES_RPTPPP, ES_UCH, ES_U, ES_AB, ES_PR))   
Set Nets = new Map
Call Nets.init(Array(NET_VO, NET_VS, NET_GS,  NET_SS, NET_TS, NET_ES), Array(VO, VS, GS,  SS, TS, ES))

'NET_DK, DK,
Dim out
Dim mapA, layer, layerName
Const ERROR_MESSAGE = "{\B\C0000FF}"
Const WARNING_MESSAGE = "{\B\C0066ff}"


Sub CheckNullData()

	init()	
	initMap()	
    
	For i = 1 to map.Layers.Count: Do        
		ExportPath = RootExportPath	
        prepareLayerName(i)
		                
        Set L = CreateObject("ZuluLib.Layer") 
		L.Open(layer.Name) 
        L.Active = True
        mess("Активный слой: {\B}" & layerName)
        	                
        Set Db = CreateObject("Zb.ZbDatabase") 
		For j = 0 To B.Count - 1  
			out.Put "{\B\C008000}  ID=" + CStr(B.Item(i).Id) + "  " + B.Item(j).UserName + CHR(10)     
			If Db.Open(B.Item(j).Name) = True Then
				field_count = Db.ActiveQuery.VisualQuery.Fields.Count
				Set Flds = Db.ActiveQuery.VisualQuery.Fields
				For k = 0 To field_count - 1
					out.Put "    " + Flds.Item(j).Name + Space(54 - Len(Flds.Item(j).Name)) + """" + Flds.Item(j).UserName + """" + CHR(10)
				Next       
			End If   
		Next  
        
    Loop While False
	Next
    
End	Sub	

Private Function prepareLayerName(index)
    
    Set layer = map.Layers.Item(index)
	mess("Обработка слоя: " & layer.Name & "...")  	
    	
	paths = Split(layer.Name, "/")
	layerName = (Split(paths(UBound(paths)), "."))(0)
    mess("Извлеченное имя слоя: " & layerName) 
End Function

Private Function mess(message) 	
   out.Put(code & message & Chr(10))   	
End Function

Private Function init()
    
    NetKeys = Array("ВС", "ТС", "ВО", "ДК")	
    Set out = Zulu.OpenOutputChannel("Сообщения") 
    out.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
	mess("Макрос проверки пустых полей в слое")
	
End Function


Private Function initMap()

	Set map = CreateObject("ZuluLib.MapDoc")	
    
	map.Open(mapName)
    mess("Название карты: " & map.Name)
End Function