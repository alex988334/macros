' ActiveX enumeration values definitions start (do not change!)
Const zbqtVisualQuery = 1
' ActiveX enumeration values definitions end

Const mapName = "zulu://10.0.2.20:6473/ТЕСТ/Test_Gorelik/Алексей_map.zmp"
Dim mapA, layer, layerName
Const ERROR_MESSAGE = "{\B\C0000FF}"
Const WARNING_MESSAGE = "{\B\C0066ff}"

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
        
		Set QryInf = db.Queries.AddNew(zbqtVisualQuery, "Структура")	
        Set VisQry = QryInf.VisualQuery

		' Добавляем таблицу в запрос
		Set TblRef = VisQry.Tables.Add(db.Tables(0), True)

		' Добавляем поле связи с картой
		VisQry.SetBaseField(TblRef, "Sys")
        Set Field = VisQry.AddField(TblRef, "Name")
    
    
End Sub


Private Function mess(message) 	
   out.Put(code & message & Chr(10))   	
End Function

Private Function init()
    
    Set out = Zulu.OpenOutputChannel("Сообщения") 
    out.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
	mess("Проверка структуры слоя...")
    	
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