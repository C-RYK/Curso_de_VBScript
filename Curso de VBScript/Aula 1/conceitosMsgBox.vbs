'MSGBOX é composto de até 5 parametros:
'	PROMPT -> que é a mensagem que aparece no centro da caixa de texto.
'	BUTTON -> que é composto de 4 grupos/números.
'		Grupo 1 -> Configurações do botão.
'		Grupo 2 -> Icones da caixa de texto.
'		Grupo 3 -> Botão padrão.
'		Grupo 4 -> Modal.
'	TITLE -> que é mensagem que aparece na barra de título da caixa de texto.
'	HELPFILE -> que é o endereço do arquivo de ajuda da caixa de texto.
'	CONTEXT -> que é o contexto do arquivo de ajuda.
'
'Grupo 1:
'	Cenario 1
'		Condições:
'			- Se informado o valor 0
'				resultado:
'					- A caixa de texto exibe o botão "Ok"
'
'	Cenario 2
'		Condições:
'			- Se informado o valor 1
'				resultado:
'					- A caixa de texto exibe o botão "Ok", "Cancelar
'
'	Cenario 3
'		Condições:
'			- Se informado o valor 2
'				resultado:
'					- A caixa de texto exibe o botão "Abortar", "Repetir" e "Cancelar"
'
'	Cenario 4
'		Condições:
'			- Se informado o valor 3
'				resultado:
'					- A caixa de texto exibe o botão "Sim", "Não" e "Cancelar"
'
'	Cenario 5
'		Condições:
'			- Se informado o valor 4
'				resultado:
'					- A caixa de texto exibe o botão "Sim" e "Não"
'
'	Cenario 6
'		Condições:
'			- Se informado o valor 5
'				resultado:
'					- A caixa de texto exibe o botão "Repetir" e "Cancelar"
'
'
'Grupo 2:
'	Cenario 1
'		condições
'			- Se informado o valor 0
'				resultado:
'					- A caixa de texto não exibirá icone
'
'	Cenario 2
'		condições
'			- Se informado o valor 16
'				resultado:
'					- A caixa de texto exibirá ícone de erro
'
'	Cenario 3
'		condições
'			- Se informado o valor 32
'				resultado:
'					- A caixa de texto exibirá ícone de questão
'
'	Cenario 4
'		condições
'			- Se informado o valor 48
'				resultado:
'					- A caixa de texto exibirá ícone de exclamação
'
'	Cenario 1
'		condições
'			- Se informado o valor 64
'				resultado:
'					- A caixa de texto exibirá ícone informativo
'
'
'Grupo 3:
'	Cenario 1
'		Condições:
'			- Se informado o valor 0
'				resultado:
'					- O primeiro botão ficará destacado.
'
'	Cenario 2
'		Condições:
'			- Se informado o valor 256
'				resultado:
'					- O Segundo botão ficará destacado.
'
'	Cenario 3
'		Condições:
'			- Se informado o valor 512
'				resultado:
'					- O Terceiro botão ficará destacado.
'
'	Cenario 4
'		Condições:
'			- Se informado o valor 768
'				resultado:
'					- O Quarto botão ficará destacado.
'
'
'Grupo 4:
'	Cenario 1
'		Condições:
'			- Se informado o valor 0
'				resultado:
'					- Modal desativado.
'
'	Cenario 1
'		Condições:
'			- Se informado o valor 4096
'				resultado:
'					- Modal ativado.


msgbox "Olá Mundo", 308, "Atenção"
