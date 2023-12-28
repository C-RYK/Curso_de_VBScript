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

' Filter tem até 4 parâmetros
'	Array - que é o array em que se faz o filtro
'	value - que é o valor a ser procurado
'	Include - que é um valor booleano (True ou False) para filtrar se contem ou não o valor do parâmetro value
'	Compare - O comparação de texto, 1 comparação binária.


b = filter(Varray, "s")

for i = 0 to ubound(b) 
	msgbox b(i)
next
