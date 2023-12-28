for cont = 1 to 10
	a = MsgBox("Repetição " & cont & " Deseja sair ?", 4)
	if a = 6 then
		Exit for
	end if
next
