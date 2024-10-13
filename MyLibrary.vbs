Private Function mess(message) 	
   out.Put(code & message & Chr(10))   	
End Function

Private Function init()
    
    Set out = Zulu.OpenOutputChannel("Сообщения") 
    out.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
	mess("Проверка структуры слоя...")
    	
End Function