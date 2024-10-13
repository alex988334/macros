Sub ReadLayerStructure

'Создаем объект для вывода в окно сообщений
Set out = Zulu.OpenOutputChannel("Сообщения")

'Очищаем окно сообщений
out.Clear

'Создаем объект Layer
Set L = CreateObject("ZuluLib.Layer")

'Открываем слой
L.Open("zulu://10.0.2.20:6473/ТЕСТ/Test_Gorelik/ТС_4_Этап_тест.zl")

out.Put "{\B}Слой: " + L.UserName + CHR(10)
	
out.Put CHR(10) + "{\B}Темы: " + CHR(10)    

'Получаем коллекцию тематических раскрасок слоя
Set Th = L.Themes

'Выводим список ID и имен тематических раскрасок
For i = 0 To Th.Count - 1    
   out.Put "  ID=" + CStr(Th.ThemeId(i)) + "  " + Th.Item(i).UserName + CHR(10) 
Next    

out.Put CHR(10) + "{\B}Надписи: " + CHR(10)    

'Получаем коллекцию вариантов надписей слоя
Set Lb = L.LabelLayers

'Выводим список ID и имен вариантов надписей слоя
For i = 0 To Lb.Count - 1
    
out.Put "  ID=" + CStr(Lb.Item(i).Id) + "  " + Lb.Item(i).UserName + CHR(10)     
    
Next

'Получаем коллекцию баз данных слоя
Set B = L.Bases

out.Put CHR(10) + "{\B}Базы данных: " + CHR(10) 

'Создаем объект ZbDatabase для доступа к базе данных
Set Db = CreateObject("Zb.ZbDatabase")       

For i = 0 To B.Count - 1
    
'Выводим ID и имя базы данных
    out.Put "{\B\C008000}  ID=" + CStr(B.Item(i).Id) + "  " + B.Item(i).UserName + CHR(10)     
    
'Открываем базу данных
    If Db.Open(B.Item(i).Name) = True Then

'Получаем количество отображаемых полей активного запроса
	field_count = Db.ActiveQuery.VisualQuery.Fields.Count

'Получаем коллекцию отбражаемых в запросе полей
        Set Flds = Db.ActiveQuery.VisualQuery.Fields
        
'Выводим список полей
        For j = 0 To field_count - 1

        out.Put "    " + Flds.Item(j).Name + Space(45 - Len(Flds.Item(j).Name)) + """" + Flds.Item(j).UserName + """" + CHR(10)
        		
        Next		
        
    End If    
    
Next    

out.Put CHR(10) + "{\B}Типы и режимы: " + CHR(10)        

'Получаем коллекцию типов слоя
Set Types = L.ObjectTypes

For i = 0 To Types.Count - 1
    
'Получаем тип по индексу
	Set ObjType = Types.GetItemByIndex(i)	
    
'Выводим IF и имя типа
    out.Put "{\B\C800000}  ID=" + CStr(ObjType.Id) + "  " + ObjType.Name + CHR(10) 
    
'Получаем коллекцию режимов типа
    Set Modes = ObjType.Modes
    
    For j = 1 To Modes.Count
            
'Выводим имя режима
        out.Put "    " + Modes.Item(j).Name + CHR(10)
          
	Next        
    
Next

End Sub