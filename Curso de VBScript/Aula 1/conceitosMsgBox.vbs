'MSGBOX � composto de at� 5 parametros:
'	PROMPT -> que � a mensagem que aparece no centro da caixa de texto.
'	BUTTON -> que � composto de 4 grupos/n�meros.
'		Grupo 1 -> Configura��es do bot�o.
'		Grupo 2 -> Icones da caixa de texto.
'		Grupo 3 -> Bot�o padr�o.
'		Grupo 4 -> Modal.
'	TITLE -> que � mensagem que aparece na barra de t�tulo da caixa de texto.
'	HELPFILE -> que � o endere�o do arquivo de ajuda da caixa de texto.
'	CONTEXT -> que � o contexto do arquivo de ajuda.
'
'Grupo 1:
'	Cenario 1
'		Condi��es:
'			- Se informado o valor 0
'				resultado:
'					- A caixa de texto exibe o bot�o "Ok"
'
'	Cenario 2
'		Condi��es:
'			- Se informado o valor 1
'				resultado:
'					- A caixa de texto exibe o bot�o "Ok", "Cancelar
'
'	Cenario 3
'		Condi��es:
'			- Se informado o valor 2
'				resultado:
'					- A caixa de texto exibe o bot�o "Abortar", "Repetir" e "Cancelar"
'
'	Cenario 4
'		Condi��es:
'			- Se informado o valor 3
'				resultado:
'					- A caixa de texto exibe o bot�o "Sim", "N�o" e "Cancelar"
'
'	Cenario 5
'		Condi��es:
'			- Se informado o valor 4
'				resultado:
'					- A caixa de texto exibe o bot�o "Sim" e "N�o"
'
'	Cenario 6
'		Condi��es:
'			- Se informado o valor 5
'				resultado:
'					- A caixa de texto exibe o bot�o "Repetir" e "Cancelar"
'
'
'Grupo 2:
'	Cenario 1
'		condi��es
'			- Se informado o valor 0
'				resultado:
'					- A caixa de texto n�o exibir� icone
'
'	Cenario 2
'		condi��es
'			- Se informado o valor 16
'				resultado:
'					- A caixa de texto exibir� �cone de erro
'
'	Cenario 3
'		condi��es
'			- Se informado o valor 32
'				resultado:
'					- A caixa de texto exibir� �cone de quest�o
'
'	Cenario 4
'		condi��es
'			- Se informado o valor 48
'				resultado:
'					- A caixa de texto exibir� �cone de exclama��o
'
'	Cenario 1
'		condi��es
'			- Se informado o valor 64
'				resultado:
'					- A caixa de texto exibir� �cone informativo
'
'
'Grupo 3:
'	Cenario 1
'		Condi��es:
'			- Se informado o valor 0
'				resultado:
'					- O primeiro bot�o ficar� destacado.
'
'	Cenario 2
'		Condi��es:
'			- Se informado o valor 256
'				resultado:
'					- O Segundo bot�o ficar� destacado.
'
'	Cenario 3
'		Condi��es:
'			- Se informado o valor 512
'				resultado:
'					- O Terceiro bot�o ficar� destacado.
'
'	Cenario 4
'		Condi��es:
'			- Se informado o valor 768
'				resultado:
'					- O Quarto bot�o ficar� destacado.
'
'
'Grupo 4:
'	Cenario 1
'		Condi��es:
'			- Se informado o valor 0
'				resultado:
'					- Modal desativado.
'
'	Cenario 1
'		Condi��es:
'			- Se informado o valor 4096
'				resultado:
'					- Modal ativado.


msgbox "Ol� Mundo", 308, "Aten��o"
