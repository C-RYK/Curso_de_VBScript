Dim varray()

for cont = 1 to 10
	redim preserve varray(cont)
	varray(cont) = "Valor " & cont
next

for each A in varray
	msgbox A
next
