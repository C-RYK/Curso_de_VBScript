set arquivoExterno = createObject("Scripting.fileSystemObject")
set arq = arquivoExterno.OpentextFile("C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Aula 25\teste.txt", 2, 1)

a = InputBox(arq.writeline("Informe o primeiro valor: "))

b = InputBox(arq.writeline("Informe o segundo valor: "))

c = cdbl(a) + cdbl(b)

MsgBox arq.writeline("A soma dos dois valores é " & c)

arq.close

msgbox "Finalizado", 64, "Arquivo Texto"
