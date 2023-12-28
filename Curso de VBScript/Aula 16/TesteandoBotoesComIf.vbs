' Cenario 1:
'	Condições:
'		- O operador escolhe a opção "Sim"
'	Resultado:
'		- Exibe a mensagem "Foi pressionado o botão Sim"

' Cenario 2:
'	Condições:
'		- O operador escolhe a opção "Não"
'	Resultado:
'		- Exibe a mensagem "Foi pressionado o botão Não"


a = msgbox("Aperte em um dos botões", 4)

if  a = 6 then
	msgbox "Foi pressionado o botão 'Sim'"

else
	msgbox "Foi pressionado o botão 'Não'"
end if

