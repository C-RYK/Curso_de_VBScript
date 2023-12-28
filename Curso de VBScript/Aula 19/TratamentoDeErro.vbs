a =  InputBox("Infome o primeiro número: ")
b =  InputBox("Infome o segundo número: ")

On error resume next
c = a / b

if err.number <> 0 then
	MsgBox "Houve um erro, tente novamente"
else
	MsgBox "O valor da divisão dos dois números é " & c
end if

on error goto 0