Const mapName = "zulu://10.0.2.20:6473/ТЕСТ/Test_Gorelik/Алексей_map.zmp"

Dim VO_SU, VO_NU, VO_K, VO_PR, VS_IV, VS_UVS, VS_U, VS_VK, VS_POT, VS_PR, GS_U, GS_UCH, GS_POT, GS_PR, DK_K, DK_SU, DK_NU, DK_PR
Dim SS_U, SS_UCH, SS_POT, SS_PR, TS_I, TS_UCH, TS_U, TS_OP, TS_PR, ES_POD, ES_RPTPPP, ES_UCH, ES_U, ES_AB, ES_PR 
 

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

Dim  NetKeys, Nets(6,1), VO(3,1), VS(5,1), GS(3,1), DK(3,1), SS(3,1), TS(4,1), ES(5,1)  

VO(0,0) = SAM_UCH	
VO(0,1) = VO_SU
VO(1,0) = NAP_UCH	
VO(1,1) = VO_NU
VO(2,0) = KOL	
VO(2,1) = VO_K
VO(3,0) = PRIM	
VO(3,1) = VO_PR

VS(0,0) = IST_VOD	
VS(0,1) = VS_IV
VS(1,0) = UCH_VOD_SET	
VS(1,1) = VS_UVS
VS(2,0) = UZEL	
VS(2,1) = VS_U
VS(3,0) = VOD_KOL_GIDRAN	
VS(3,1) = VS_VK
VS(4,0) = POT	
VS(4,1) = VS_POT 
VS(5,0) = PRIM	
VS(5,1) = VS_PR

GS(0,0) = UZEL	
GS(0,1) = GS_U
GS(1,0) = UCH	
GS(1,1) = GS_UCH
GS(2,0) = POT	
GS(2,1) = GS_POT
GS(3,0) = PRIM	
GS(3,1) = GS_PR

DK(0,0) = KOL	
DK(0,1) = DK_K
DK(1,0) = SAM_UCH	
DK(1,1) = DK_SU
DK(2,0) = NAP_UCH	
DK(2,1) = DK_NU
DK(3,0) = PRIM	
DK(3,1) = DK_PR

SS(0,0) = UZEL	
SS(0,1) = SS_U	
SS(1,0) = UCH	
SS(1,1) = SS_UCH
SS(2,0) = POT	
SS(2,1) = SS_POT
SS(3,0) = PRIM	
SS(3,1) = SS_PR

TS(0,0) = IST	
TS(0,1) =  TS_I
TS(1,0) =  UCH	
TS(1,1) =  TS_UCH
TS(2,0) =  UZEL	
TS(2,1) =  TS_U
TS(3,0) =  OBOB_POT	
TS(3,1) =  TS_OP
TS(4,0) =  PRIM	
TS(4,1) =  TS_PR

ES(0,0) = PODST	
ES(0,1) = ES_POD
ES(1,0) = RP_TP_PP	
ES(1,1) = ES_RPTPPP
ES(2,0) = UCH	
ES(2,1) = ES_UCH
ES(3,0) = UZEL	
ES(3,1) = ES_U
ES(4,0) = ABON	
ES(4,1) = ES_AB
ES(5,0) = PRIM	
ES(5,1) = ES_PR

Nets(0,0) = NET_VO
Nets(0,1) = VO
Nets(1,0) = NET_VS
Nets(1,1) = VS
Nets(2,0) = NET_GS
Nets(2,1) = GS
Nets(3,0) = NET_SS
Nets(3,1) = SS
Nets(4,0) = NET_TS
Nets(4,1) = TS
Nets(5,0) = NET_ES
Nets(5,1) = ES
Nets(6,0) = NET_DK
Nets(6,1) = DK

NetKeys = Array(NET_VO, NET_VS, NET_GS, NET_SS, NET_TS, NET_ES, NET_DK)
		
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

Dim out
Dim mapA, layer, layerName
Const ERROR_MESSAGE = "{\B\C0000FF}"
Const WARNING_MESSAGE = "{\B\C0066ff}"


Sub checkLayerStructure

	init()	
	initMap() 

	For i = 1 to mapA.Layers.Count: Do        
		ExportPath = RootExportPath	
        prepareLayerName(i)
		                
        Set L = CreateObject("ZuluLib.Layer") 
		L.Open(layer.Name) 
        L.Active = True
        netName = findNet()
        If netName = Null Then		
			mess(ERROR_MESSAGE & "ERROR!!! В названии слоя не найден тип сети, имя слоя: '" & layerName & "'")	
			Exit Do	
		End If
        
        Set B = L.Bases		
        mess("Сеть: " & netName)
                
        n = getNet(Nets, netName)       
        
		Set B = L.Bases
		'out.Put CHR(10) + "{\B}Базы данных: " + CHR(10) 
		Set Db = CreateObject("Zb.ZbDatabase")     
        basesNames = ""
        r = 1
        For j = 0 To B.Count - 1 
            If (Len(basesNames) - (r * 125)) > 0 Then                            
				basesNames = basesNames & Chr(10)
				r = r + 1	
			End If
            basesNames = basesNames & "'" & B.Item(j).UserName & "', "	
		Next
        
		For k = 0 To UBound(n): do
            typen = n(k, 0)
            flag = true
            Dim base
            For j = 0 To B.Count - 1 : 
                If StrComp(typen, B.Item(j).UserName, 1) = 0 then
                    flag = false
                    Set base = B.Item(j)                    
                    mess("{\B}Проверяется бд: '" & base.UserName & "'")
                    Exit for	
				End If
			Next
            If flag then                 
				mess(ERROR_MESSAGE & "Не найдена база данных для типа: '" & typen & "'")
                mess("{\B}Список баз данных: {\B}" & Chr(10) & basesNames)
                Exit Do	
            End If
            
            printFields = false            
            attrs = getAttr(n, typen)                    
            For t = 0 to UBound(attrs)
                flag = true
				If Db.Open(base.Name) = True Then
					field_count = Db.ActiveQuery.VisualQuery.Fields.Count
					Set Flds = Db.ActiveQuery.VisualQuery.Fields
					For j = 0 To field_count - 1                       
                        If StrComp(attrs(t), Flds.Item(j).Name, 1) = 0	Then
							flag = false
                        End If
					Next
					If flag then
						mess(ERROR_MESSAGE & "Отсутствует поле: '" & attrs(t) & "'")
                        printFields = true	
                    End if
				Else 
					mess(ERROR_MESSAGE & "Отсутствует соединение с бд: '" & B.Item(k).Name & "'")
				End If
            Next
            If printFields Then
                If Db.Open(base.Name) = True Then
					field_count = Db.ActiveQuery.VisualQuery.Fields.Count
					Set Flds = Db.ActiveQuery.VisualQuery.Fields
                    mess("{\B}Существующие поля: ")
                    fields = ""
                    r = 1
					For j = 0 To field_count - 1                          
                        If (Len(fields) - (r * 125)) > 0 Then                            
                            fields = fields & Chr(10)
                            r = r + 1	
						end If
                        fields = fields & "'" & Flds.Item(j).Name & "', "                        
					Next
                    mess(fields)						
				Else 
					mess(ERROR_MESSAGE & "Отсутствует соединение с бд: '" & B.Item(k).Name & "'")
				End If
			End if
        Loop While False    
		Next
	Loop While False   
	Next
End sub 

Private function getAttr(ByRef types, TypeName)

	For i = 0 to UBound(types)
		If StrComp(types(i, 0), TypeName, 1) = 0 Then
            getAttr = types(i, 1)
            Exit function	
		End If
    Next
	getAttr = null
End Function

Private function getNet(ByRef nets, netName)
	    
    For i = 0 to UBound(nets)
		If StrComp(nets(i, 0), netname, 1) = 0 Then
            getNet = nets(i, 1)
            Exit function	
		End If
    Next
	getNet = null
End function		
       
  

Private Function mess(message) 	
   out.Put(code & message & Chr(10))   	
End Function



Private Function findNet()
	
    net = Null
    ln = "_" & layerName
	For i = 0 To UBound(NetKeys) 
		If (InStr(ln, NetKeys(i)) > 0) Then  			
			net = NetKeys(i)
            Exit for	
		End If			
	Next
    findNet = net	
End function

Private Function init()
    
    Set out = Zulu.OpenOutputChannel("Сообщения") 
    out.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
	mess("Проверка структуры слоя...")
    	
End Function

Private Function initMap()

	Set mapA = CreateObject("ZuluLib.MapDoc")	
    
	mapA.Open(mapName)
    mess("Название карты: " & mapA.Name)
End Function

Private Function prepareLayerName(index)
    
    Set layer = mapA.Layers.Item(index)
    
	mess(CHR(10) & "{\B}Обработка слоя: " & layer.Name)  	
    	
	paths = Split(layer.Name, "/")
	layerName = (Split(paths(UBound(paths)), "."))(0)
   ' mess("Извлеченное имя слоя: " & layerName) 
End Function

