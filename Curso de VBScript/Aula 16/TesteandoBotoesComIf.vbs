' Cenario 1:
'	Condi��es:
'		- O operador escolhe a op��o "Sim"
'	Resultado:
'		- Exibe a mensagem "Foi pressionado o bot�o Sim"

' Cenario 2:
'	Condi��es:
'		- O operador escolhe a op��o "N�o"
'	Resultado:
'		- Exibe a mensagem "Foi pressionado o bot�o N�o"


a = msgbox("Aperte em um dos bot�es", 4)

if  a = 6 then
	msgbox "Foi pressionado o bot�o 'Sim'"

else
	msgbox "Foi pressionado o bot�o 'N�o'"
end if

