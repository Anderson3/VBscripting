' ScriptCryptor Project Options Begin
' HasVersionInfo: No
' Companyname: 
' Productname: 
' Filedescription: 
' Copyrights: 
' Trademarks: 
' Originalname: 
' Comments: 
' Productversion:  0. 0. 0. 0
' Fileversion:  0. 0. 0. 0
' Internalname: 
' Appicon: 
' AdministratorManifest: No
' ScriptCryptor Project Options End


dim hora, minuto, mensagem

hora = cint(inputBox("Defina a HORA de inicialização da fala: "))
minuto = cint(inputBox("Defina o MINUTO de inicialização da fala: "))

do
	if Hour(Now) = hora AND Minute(Now) = minuto then
		set tts = createObject("sapi.spvoice")
		
		set fso = createObject("scripting.fileSystemObject")
		const ForReading = 1, ForWriting = 2, ForAppending = 8
		set leitor = fso.openTextFile("texto.txt", ForReading)
		mensagem = leitor.readAll
		leitor.close
			
		tts.speak mensagem
	end if

		wscript.sleep (32*1000)
loop
