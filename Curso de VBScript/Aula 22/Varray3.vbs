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

'msgbox "A variavel contador é array: " & isArray(contador)
'msgbox "A variavel Varray é array: " & isArray(Varray)

'msgbox "O resultado da função join(Varray) é array: " & isArray(join(Varray))

'Join - função que transforma um array em um valor caracterer
'	Array - Que o array que será transformado em texto
'	Delimiter - Que é um delimitador dos valores do array. Default " "

texto = join(Varray, ",")

'Slipt - transforma um texto em array
'	value - Que é o texto a ser transformado em array
'	Delimitador - Que é o caracter(es) que delitam os valores a serem transformados
'	count - Que define a quantidade de substrings a serem retornadas
'	compare - Que é tipo de comparação 0 - Binaria e 1 - Textual

Varray2 = Split(texto, "###", 4)

'msgbox "A variavel Varray2 é array: " & isArray(Varray2)

for each b in Varray2
	msgbox b
next
