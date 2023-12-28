set ArquivoExterno = createObject("Scripting.FileSystemObject")
textoFinal = ""

Set arqLeitura = ArquivoExterno.GetFile("C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Aula 26\teste.txt")
Set arqStream = arqLeitura.OpenAsTextStream(1, -2)

Do while not arqStream.AtEndOfStream
	sRecord = arqStream.readLine
	textoFinal  = textoFinal & sRecord & chr(10)
loop

arqStream.close
msgBox textoFinal
