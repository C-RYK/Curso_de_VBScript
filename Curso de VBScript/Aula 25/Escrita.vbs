set arquivoExterno = createObject("Scripting.fileSystemObject")
set arq = arquivoExterno.OpentextFile("C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Aula 25\teste.txt", 2, 1)

arq.writeline "Primeira linha"
arq.writeline "Segunda linha"
arq.writeline "Terceira linha"
arq.writeline "Quarta linha"
arq.writeline "Quinta linha"
arq.writeline "Sexta linha"
arq.writeline "Sétima linha"

arq.close

msgbox "Finalizado", 64, "Arquivo Texto"
