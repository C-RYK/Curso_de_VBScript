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

' Filter tem at� 4 par�metros
'	Array - que � o array em que se faz o filtro
'	value - que � o valor a ser procurado
'	Include - que � um valor booleano (True ou False) para filtrar se contem ou n�o o valor do par�metro value
'	Compare - O compara��o de texto, 1 compara��o bin�ria.


b = filter(Varray, "s")

for i = 0 to ubound(b) 
	msgbox b(i)
next
