Set arquivo = CreateObject("Scripting.FileSystemObject")
set arq = arquivo.GetFile("C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Aula 27\teste.txt")

propriedades = "Data da criação - ["& arq.DateCreated &"]" & chr(10) & _
		      "Data do último acesso - ["& arq.DateLastAccessed &"]" & chr(10) & _
                          "Data da última modificação - ["& arq.DateLastModified &"]" & chr(10) & _
		      "Drive - ["& arq.drive &"]" & chr(10) & _
		      "Nome - ["& arq.name &"]" & chr(10) & _
		      "Pasta - ["& arq.parentFolder &"]" & chr(10) & _
		      "Nome Completo - ["& arq.path &"]" & chr(10) & _
		      "Nome Dos - ["& arq.ShortName &"]" & chr(10) & _
		      "Nome Completo no DOS - ["& arq.ShortPath &"]" & chr(10) & _
		      "Tamanho - ["& arq.size &"]" & chr(10) & _
		      "Extensão - ["& arquivo.getExtensionName(arq.path) &"]" & chr(10) & _
		      "Tipo - ["& arq.type &"]"

msgbox propriedades


set rede = createObject("wScript.Network")
msgbox "Você está conectado com o usúario: " & rede.username
