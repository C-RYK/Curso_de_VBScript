a = InputBox ("Qual das opções você quer criar? " & chr(10) & _
	      "(Escolha o número 1 ou 2) " & chr(10) & _
	      chr(10) & _
	      " 1 - Arquivo Txt" & chr(10) & _
	      " 2 - Arquivo Vbs")

if a = 1 then
	set arquivoExterno = createObject("Scripting.fileSystemObject")
	set arq = arquivoExterno.OpentextFile("C:\Users\carlo\OneDrive\Documentos\Praticando Vbs\Praticas.txt", 2, 1)

	arq.writeline "Primeira linha"
	arq.writeline "Segunda linha"
	arq.writeline "Terceira linha"
	arq.writeline "Quarta linha"
	arq.writeline "Quinta linha"
	arq.writeline "Sexta linha"
	arq.writeline "Sétima linha"
	
	arq.close

elseif a = 2 then
	set arquivoExterno = createObject("Scripting.fileSystemObject")
	set arq = arquivoExterno.OpentextFile("C:\Users\carlo\OneDrive\Documentos\Praticando Vbs\Praticas.vbs", 2, 1)
	
	a = InputBox("Informe o primeiro valor: ")
	b = InputBox("Informe o segundo valor: ")
	c = cdbl(a) + cdbl(b)
	Msgbox "O resultado da soma de " & a & " + " & b & " é " & c
	
	arq.close
end if


