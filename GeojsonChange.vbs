executeGlobal(CreateObject("Scripting.FileSystemObject").openTextFile(fname).readAll())


Set out = Zulu.OpenOutputChannel("Сообщения")
Const FilePath = "O:\Export\geojson\"
Const FileName = "primer1.geojson"
Dim findAttr, delimeter
delimeter = "'"
findAttr = Array("typename", "modename")
  
Class Map
	Private ke
    Private va
    
    Function init()        
        ke = Array("Источник", "Обобщенный потребитель", "Участки", "Узел")
        va = Array(",'icon':'ist.icon','iconWidth':64,'iconHeight':64,", ",'icon':'potr.icon','iconWidth':64,'iconHeight':64,", 
        ",'icon':'uch.icon','iconWidth':64,'iconHeight':64,", ",'icon$:$uzel.icon','iconWidth':64,'iconHeight':64,")
    End Function
            
    Function GetValue(index)    
        If index = NULL Then 
            GetValue = ""
            Exit function	
		End If
        For k = 0 To UBound(ke) 	           			If (StrComp(index, ke(k)) = 0) Then 
                GetValue = va(k)
                Exit function		
			End If
		Next        
        GetValue = ""
	End Function  
End Class	
	
    
Sub GeojsonChange    
    out.Clear
    mess("GeojsonChange...")	    
    
    Set m = new Map    
    m.init()   
    
    Set FSO = CreateObject("Scripting.FileSystemObject")    
    Set f = FSO.OpenTextFile(FilePath & FileName)
    
	Do While Not f.AtEndOfStream
		str = f.ReadLine	
        mess(str)
        mass = Split(str, ",")
        For	j = 0 To UBound(mass) - 1
            sear1 = InStr(mass(j), findAttr(0))
            sear2 = InStr(mass(j))
            If (sear1 <> "") And (sear1 > 0) Then                ma = Split(mass(j), ":")
                val = m.GetValue(Replace(ma(1), "'", ""))
                If val <> "" Then
                    
				End If                
			End If
		Next	
        'mess(m.GetValue("Участки"))			 
	Loop
    
	f.Close
    
End Sub 


Function mess(message) 
	
    out.Put(message & Chr(10))
End Function


