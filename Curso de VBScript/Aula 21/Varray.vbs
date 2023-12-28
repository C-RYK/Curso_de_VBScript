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

for i = 1 to ubound(Varray) - 1
	msgbox Varray(i)
next
