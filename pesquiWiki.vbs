Dim nome, pesquisa,opcao,pergunta

pergunta = MsgBox("Olá, sou o pesquiWiki, gostaria de me executar agora?",4)


If pergunta = 6 Then
	
	nome = InputBox("Olá, para nos conhecermos melhor me diga seu nome!")

	pesquisa = InputBox("Olá "& nome &", bem vindo ao pesquiWiki!"& Chr(10) & _ 
	"Minha função é pesquisar conteúdos no site do wikipédia." & Chr(10) & _
	"Então vamos começar!"& Chr(10) & _ 
	"Digite o que deseja pesquisar!","Sou pesquiWiki" )
	
	If pesquisa = "" Then
		opcao = 7
		MsgBox "Que pena, você não pesquisou nada!"
	Else
		opcao = 8
	End If
	
	Do While opcao <> 7:

	Set openGoogle = CreateObject("wscript.shell")
	openGoogle.Run "https://pt.wikipedia.org/wiki/Wikip%C3%A9dia:P%C3%A1gina_principal"

	WScript.Sleep 2000
	openGoogle.SendKeys "{f3}"
	WScript.Sleep 1000
	openGoogle.SendKeys "Pesquisar"
	WScript.Sleep 1000
	openGoogle.SendKeys "{esc}"
	WScript.Sleep 1000
	openGoogle.SendKeys "{esc}"
	WScript.Sleep 1000
	openGoogle.SendKeys pesquisa
	WScript.Sleep 1000
	openGoogle.SendKeys "{enter}"
	WScript.Sleep 5000
	openGoogle.SendKeys "^{a}"
	WScript.Sleep 4000
	openGoogle.SendKeys "^{c}"


	Set openWord = CreateObject("wscript.shell")
	openWord.Run "WINWORD"
	WScript.Sleep 4000
	openWord.SendKeys "{enter}"
	openWord.SendKeys "^{v}"
	WScript.Sleep 15000
	openWord.SendKeys "^{s}"
	WScript.Sleep 1000
	openWord.SendKeys "{tab}"
	WScript.Sleep 1000
	openWord.SendKeys "{down}"
	WScript.Sleep 1000
	openWord.SendKeys "{down}"
	WScript.Sleep 1000
	openWord.SendKeys "{enter}"
	WScript.Sleep 1000
	openWord.SendKeys "{tab}"
	WScript.Sleep 1000
	openWord.SendKeys "{down}"
	WScript.Sleep 1000
	openWord.SendKeys pesquisa
	WScript.Sleep 1000
	openWord.SendKeys "{tab}"
	WScript.Sleep 1000
	openWord.SendKeys "{tab}"
	WScript.Sleep 1000
	openWord.SendKeys "{enter}"
	WScript.Sleep 6000
	
	opcao = MsgBox("deseja pesquisar novamente?", 4)
	If opcao = 7 Then
	MsgBox "obrigado!"
	End If
	WScript.Sleep 1000


	strComputer = "."

	Set objNetwork = CreateObject("Wscript.Network")

	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	''' Processo que será verificado '''''''
	Set colProcesses = objWMIService.ExecQuery _
		("Select * from Win32_Process Where Name = 'WINWORD.exe'")

	''' elimina o processo definido '''
	For each Processo in ColProcesses
	   Processo.Terminate()
	Next

	WScript.Sleep 1000
	openWord.SendKeys "^{w}"

	Loop

ElseIf pergunta = 7 Then
	MsgBox "obrigado!"

End If