Dim Varray()

contador = 1

On error resume next

Do While True
	redim preserve Varray(contador)
	Varray(contador) = weekdayName(contador)
	if err.number <> 0 then
		exit do
	end if

	contador = contador + 1
loop
On error goto 0

'msgbox "A variavel contador � array: " & isArray(contador)
'msgbox "A variavel Varray � array: " & isArray(Varray)

'msgbox "O resultado da fun��o join(Varray) � array: " & isArray(join(Varray))

'Join - fun��o que transforma um array em um valor caracterer
'	Array - Que o array que ser� transformado em texto
'	Delimiter - Que � um delimitador dos valores do array. Default " "

texto = join(Varray, ",")

'Slipt - transforma um texto em array
'	value - Que � o texto a ser transformado em array
'	Delimitador - Que � o caracter(es) que delitam os valores a serem transformados
'	count - Que define a quantidade de substrings a serem retornadas
'	compare - Que � tipo de compara��o 0 - Binaria e 1 - Textual

Varray2 = Split(texto, "###", 4)

'msgbox "A variavel Varray2 � array: " & isArray(Varray2)

for each b in Varray2
	msgbox b
next
