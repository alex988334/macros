Public Const xlToRight = -4161
Public Const xlDown = -4121
Public Const xlSrcRange = 1
Public Const xlYes = 1
Public Const xlDatabase = 1
Public Const xlOpenXMLWorkbookMacroEnabled = 52
Dim exportedTypes
Dim exportMeropriyatIndex
Dim selectedFieldsQuery
Dim associatDYColumns
Dim associatDYValues
Dim out

' Функция настройки пользовательских параметров.
Function initUsersSettings()     
    
    ' Массив хранит список экспортируемых типов
	exportedTypes = Array("Участки", "Узел", "Обобщенный потребитель")	
    
    ' В массиве хранятся номера столбцов каждого типа, которые являются полями мероприятий
    exportMeropriyatIndex = Array(2, 3, 1)    
    
    ' В массиве хранятся списки полей sql запроса, по одному для каждого типа
    selectedFieldsQuery = Array(Array("Sys", "Этап", "Назначение участка", "Вид сети", "Длина участка, м", _
				"Внутpенний диаметp подающего тpубопpовода, м", "Внутренний диаметр обратного трубопровода, м", _
				"Внутpенний диаметp подающего тpубопpовода (ГВС), м", _
				"Внутренний диаметр обратного трубопровода (ГВС), м", "V_Dpod_dla_kart"), _
        Array("Sys", "Отношение к зоне КСИО", "Этап", "Назначение камеры", "Наименование узла"), _
        Array("Sys", "Мероприятия по ЦТП/ИТП", "V_KSIO", "Этап", "Наименование узла", "Name", "CTP_ITP_Name") )	
    
End Function



Sub ExtructDataToExcel
	
    initUsersSettings
    Set out = Zulu.OpenOutputChannel("Сообщения")
    out.Clear
       	    
    Set map = Zulu.ActiveMapDoc	    
    sys = -1
	
		Set borderlayer = map.Layers.Active
          
		For k = 1 To borderlayer.ElementKeys.Count		
			Dim key 
			key = borderlayer.ElementKeys.Item(k)                
              
			Set elem = borderlayer.Elements.GetElement(key)
              
			If elem.GraphType = 5 Then
                                  
				If elem.Key > sys Then					
					sys = elem.Key
				End if
			End if
		Next	
	
    out.Put(sys & Chr(10))	
    If sys = -1 Then 
        out.Put("Не найден полигон выборки, работа прекращена"  & Chr(10))
        exit sub
	End If
    
    
        
    Set objXL = CreateObject("Excel.Application")
    objXL.Visible = TRUE
    objXL.Application.DisplayAlerts = False   
    
	Set wBook = objXL.workbooks.add()	
    wBook.Activate
    
    For i = 0 To UBound(exportedTypes)
        
        dim typeName
        typeName = exportedTypes(i)
        out.Put(TypeName  & Chr(10))	        	
        
        Dim sheet
		If i = 0 Then
            Set sheet = wBook.Sheets(1)
		Else
			Set sheet = wBook.Sheets.Add
        End if
        
        
        If TypeName = "Обобщенный потребитель" Then 
            sheet.Name = "Обобщенный_потребитель"  
		Else
			sheet.Name = TypeName
		End if
        sheet.Activate
        
        out.Put("sheet.Name => " & sheet.Name & Chr(10))
        
        line = 1
        Do 
			For k = 1 To map.Layers.Count				               
				Set lay = map.Layers.Item(k)                
                If lay.Visible = true then                    
                    exit do
				End if
			Next
            
            Exit Sub
            
        Loop while false	
		Dim selFields
        selFields = generateSelect(selectedFieldsQuery(i))
           
        For k = 1 To map.Layers.Count	
            Do                
				Set lay = map.Layers.Item(k)
                
                If lay.Visible = false then
                    exit do
				end if
                
				If lay.UserName = borderlayer.UserName then
					exit do
				End if
                
                If k = 1 Then
					For m = 0 To UBound(selectedFieldsQuery(i))	
                        val = selectedFieldsQuery(i)(m)
                        If val = "Обобщенный потребитель" then
                            val = "Обобщенный_потребитель"
						End if                            
                        
						sheet.Cells(line, m + 1).Value = val
                        
                        
                        If m = UBound(selectedFieldsQuery(i)) Then
                            sheet.Cells(line, m + 2).Value = "Мероприятие"
                            if typeName = "Участки" Then
								sheet.Cells(line, m + 3).Value = "2Ду"
                            End If
						End if
					Next
                    line = line + 1
				End if  
                
                query = "SELECT " & selFields & " FROM [" & _
						lay.UserName & "] AS L1, [" & borderlayer.UserName & "] AS L2 WHERE L2.sys = " & _
                        sys & " AND L1.Geometry.STIntersects(L2.Geometry) AND L1.typename='" & TypeName + "'" & _
                        " AND L1.[Этап] = '" & extractEtap(lay.UserName) & "'"

                out.Put(query & Chr(10))	
                
                Set rezult = map.ExecSQL(query)
                
                if rezult.RetCode <> 0 Then
					out.Put("Ошибка SQL! " & rezult.ErrorString)
				End if
                
                If rezult.DataSet.MoveFirst then   
                    currentMeropriyat = ""                 
					Do 
                        For f = 0 To rezult.DataSet.FieldCount - 1				
							dataVal = rezult.DataSet.FieldValue(f)    
                            
                       
							
                            sheet.Cells(line, f + 1).Value = dataVal                 
                 
							If f = exportMeropriyatIndex(i) Then 
                                currentMeropriyat = extructMeropriyatie(dataVal)
							End if 
							If f = UBound(selectedFieldsQuery(i))  Then
                                sheet.Cells(line, f + 2).Value = currentMeropriyat
                                if TypeName = "Участки" Then
									sheet.Cells(line, f + 3).FormulaR1C1 = "=CONCATENATE(""2Ду"",RC[-2],IF(RC[-4]*1000 <> 0, CONCATENATE(""/"",RC[-4]*1000), """"), IF(RC[-3]*1000 <> 0, CONCATENATE(""/"",RC[-3]*1000), """"))"
								End if
                            End if
						Next
                        line = line + 1
                        out.Put("line => " & line & chr(10))
					Loop while rezult.DataSet.MoveNext	
                                  
				end if                
			Loop While false
		Next 
        wBook.Saveas "O:\Example.xlsm", xlOpenXMLWorkbookMacroEnabled
	Next
     
    objXL.CutCopyMode = False    
    
    wBook.Saveas "O:\Example.xlsm", xlOpenXMLWorkbookMacroEnabled   
    
    out.Put("Завершено!")
End Sub


Function extructMeropriyatie(merop)
	
    If  InStr(1, merop, "демон", 1) > 0 AND InStr(1, merop, "байп", 1) > 0 Then
        extructMeropriyatie = "Демонтаж байпаса"
        exit function
	End if
    
    If InStr(1, merop, "демон", 1) > 0 Then
        extructMeropriyatie = "Демонтаж"
        exit function
	End if	
    
    If InStr(1, merop, "рекон", 1) > 0 Then
        extructMeropriyatie = "Реконструкция"
        exit function
	End if	
    
    If  InStr(1, merop, "строит", 1) > 0 AND InStr(1, merop, "байп", 1) > 0 Then
        extructMeropriyatie = "Строительство байпаса"
        exit function
	End if	
    
    If InStr(1, merop, "строит", 1) > 0 Or InStr(1, merop, "нов", 1) > 0 Then
        extructMeropriyatie = "Строительство"
        exit function
	End if	
    
End Function

Function extractEtap(layerUserName)
    
	lines = Split(layerUserName, "_")
    firstNumber = ""
    secondNumber = ""
    thirdNumber = ""
     
    For i = 0 To UBound(lines)
        
        If InStr(1, lines(i), "эт", 1) > 0 Then
            firstNumber = Replace(lines(i), "эт", "")	
		End if
        
        If InStr(1, lines(i), "пдт", 1) > 0 Then
            secondNumber = Replace(lines(i), "пдт", "")	
            If Len(secondNumber) = 2 Then
                thirdNumber = Mid(secondNumber, 2, 1)	
                secondNumber = Mid(secondNumber, 1, 1)
			End if            
		End if
    Next	
    
    str = firstNumber & "_Этап " & secondNumber
    If thirdNumber <> "" Then
		str = str & " подэтап " &  thirdNumber
	End If
    extractEtap = str
End Function

Function generateSelect(fieldsArr)
	
    str = ""       
    For i = 0 To UBound(fieldsArr)
        str = str & "L1.[" & fieldsArr(i) & "]"
        If i < UBound(fieldsArr) Then 
            str = str & ", "        
        End if    
	Next
    generateSelect = str
End Function

