'	Управляющие параметры
Const flagTypePath = True
Const flagModePath = True
Const flagEtapPath = False
Const mapName = "zulu://10.0.2.20:6473/ТЕСТ/Test_Gorelik/Алексей_map.zmp"
Dim RootExportPath
Const delimeterLayerName = "_"
Dim NetKeys
RootExportPath = "O:\Export\"

'	Системные переменные и константы
Const ERROR_MESSAGE = "{\B\C0000FF}"
Const WARNING_MESSAGE = "{\B\C0066ff}"
Dim out,fso
Dim ExportPath, NetPath, TypePath, ModePath
Dim map, layer, layerName
Dim net, etap, typen, moden
Dim countObT, countObM, countObL, totalCopyObj

Sub Sortirovshik
    
	init()	
	initMap()	
    
	For i = 1 to map.Layers.Count: Do        
		ExportPath = RootExportPath	
        prepareLayerName(i)
		                
        Set L = CreateObject("ZuluLib.Layer") 
		L.Open(layer.Name) 
        L.Active = True
        	                
        If countObjectsInLayer(L) = 0 Then
            Exit Do
		End If
        
        If generatePathFromLayerName() = False then
			Exit Do
		End If
		
		For k = 0 To L.ObjectTypes.Count - 1 : Do			
            Set ObjType = L.ObjectTypes.GetItemByIndex(k)   
            typen = clean(ObjType.Name)   
                       
            If countObjectsInType(L, ObjType.Id) = 0 Then                                Exit Do
			End If
				
			TypePath = ExportPath & "\" & typen
            
            For j = 1 To ObjType.Modes.Count: Do
                Set Mode = ObjType.Modes.Item(j) 
                moden = clean(Mode.Name)
                                
                If countObjectsInMode(L, ObjType.Id, Mode.Id) = 0 Then					
					Exit Do
				End If
                
                ModePath = TypePath & "\" & moden 
                newFolder(ModePath)
			                
             	
               ' rez = selectObjects()
              '  If L.Selection.ElementKeys.Count = 0 Then
               '     mess(WARNING_MESSAGE & "Режим " & moden & " типа " & typen & " пуст!")
              '      Exit Do
				'End If
                mess("Всего копируемых элементов: " & L.Selection.ElementKeys.Count)
               ' Set col = L.SelectByType(ObjType.Id, Modes.Item(j).Id)                
                newLayName = net & "_" & ObjType.Name & "_" & Mode.Name
                
			    mess(newLayName)   
                         
                If L.CopyLayer(ModePath, newLayName, 80000000) Then
                    mess("копирование - успех")
				End if
                
          '      Set copLayer = CreateObject("ZuluLib.Layer")
			'	copLayer.Open(ModePath & "\" & newLayName & ".zl")
				
                Exit For
			Next  
            Loop While False
            Exit For     
		Next
        mess("Слой подразбит: " & layer.Name)
        Loop While False
	Loop While False
	Exit for	
	Next
    
	Set fso = nothing  
End Sub


Private function readEtap()

End Function


Private function getValues()

Private function getDistinctValues(table, fields, keys, params)

	query = "SELECT DISTINCT"
End Function





Private	function countObjectsInLayer(ByRef context)
	
	Call selectObjects(context, Null, Null) 
    countObL = context.Selection.ElementKeys.Count
    if countObL = 0 Then
		mess(WARNING_MESSAGE & "Слой пустой: " & layerName) 
    End If 
    
    countObjectsInLayer = countObL
End function 


Private function countObjectsInType(ByRef context, typeid)
	
    Call selectObjects(context, Array("typeid"), Array(typeid)) 
    countObT = context.Selection.ElementKeys.Count
    if countObT = 0 Then
		mess(WARNING_MESSAGE & "Тип пустой: слой - " & layerName & ", тип - " & typen) 
    End If 
    
    countObjectsInType = countObT
End function


Private function countObjectsInMode(ByRef context, typeid, modeid)
	
    Call selectObjects(context, Array("typeid", "modeid"), Array(typeid, modeid)) 
    countObM = context.Selection.ElementKeys.Count
    if countObM = 0 Then
		mess(WARNING_MESSAGE & "Режим пустой: слой - " & layerName & ", тип - " & typen & ", режим - " & moden)
    End If
    
    countObjectsInMode = countObM
End Function	


Private Function selectObjects(ByRef context, ByRef keyParams, ByRef params)

	query = "ALTER SELECTION ON [" & layerName & "] ADD SELECT Sys FROM [" & layerName & "]"
    
    If Not ExecSQL(context, query & generateWhere(keyParams, params)) Then
        selectObjects = false
        Exit Function		
    End If   
    
    selectObjects = True
End Function


Private function generateSelect(fields, distinct)

	sel = "SELECT "
    
    If fields = null Then		sel = sel & "*"
	Else 
    
    End if
End function


Private Function generateWhere(ByRef keyParams, ByRef params)

	If keyParams = Null OR params = Null Then
		mess(WARNING_MESSAGE & " Параметры или ключи SQL равны Null!")
        generateWhere = ""
        Exit function	
	End If
    
    If UBound(keyParams) <>	UBound(params) Then
		mess(ERROR_MESSAGE & " Несоответствие параметров SQL и их ключей! UBound(keyParams): " & UBound(keyParams) & ", UBound(params): " & UBound(params))		
		generateWhere = ""
		Exit Function	
    End If 
           	
	where = " WHERE"
	For i = 0 to UBound(keyParams)	
		and_ = " AND" 
		If i = 0 Then 
			and_ = "" 
		End If
		where = where & and_ & " [" & keyParams(i) & "] = '" & params(i) & "'"
	Next 
           
	generateWhere = where
End Function


Private Function ExecSQL(context, query)
	
    Set rez = context.ExecSQL(query)
    If rez.RetCode = 0 Then		
        ExecSQL = True	
	Else
		mess(ERROR_MESSAGE & " Ошибка SQL запроса: " & rez.ErrorString)		ExecSQL = False	
    End If    
End Function


Private function clean(str)
	clean = Trim(Replace(str, Chr(34), ""))		
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

' не реализована
Private Function findEtap()

	etap = Null
    findEtap = False	
End Function


Private Function generatePathFromLayerName()

	splitN = Split(layerName, delimeterLayerName)
    If 	findNet() = Null Then		
        mess(ERROR_MESSAGE & "ERROR!!! В названии слоя не найден тип сети, имя слоя: " & layerName)	
        generatePathFromLayerName = False
        mess("конец")		
		Exit Function	
	End If
    
    ExportPath = ExportPath & "\" & net		
    
    If flagEtapPath Then        
        If findEtap() = False Then
            mess(WARNING_MESSAGE & "WARNING! В названии слоя этап не найден, имя слоя: " & layerName)	           
        Else			ExportPath = ExportPath & "\" & etap	
        End If 
	End If
    
    newFolder(ExportPath)	    
    generatePathFromLayerName = True    
End Function


Private Function prepareLayerName(index)
    
    Set layer = map.Layers.Item(index)
	mess("Обработка слоя: " & layer.Name & "...")  	
    	
	paths = Split(layer.Name, "/")
	layerName = (Split(paths(UBound(paths)), "."))(0)
    mess("Извлеченное имя слоя: " & layerName) 
End Function


Private Function init()
    
    NetKeys = Array("ВС", "ТС", "ВО", "ДК")	
    Set out = Zulu.OpenOutputChannel("Сообщения") 
    out.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
	mess("Сортировщик...")
	
    RootExportPath = RootExportPath & Replace(Replace(Now(), ":", ""), " ", "_")    
    newFolder(ExportPath)		
End Function


Private Function initMap()

	Set map = CreateObject("ZuluLib.MapDoc")	
    
	map.Open(mapName)
    mess("Название карты: " & map.Name)
End Function

    
Private Function newFolder(path)
    mess("Create path: " & path)		
	If fso.FolderExists(path) = false Then
		fso.CreateFolder (path)
	End If	  
End Function
    
    
Private Function mess(message) 	
   out.Put(code & message & Chr(10))   	
End Function