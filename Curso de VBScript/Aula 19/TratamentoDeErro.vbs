a =  InputBox("Infome o primeiro n�mero: ")
b =  InputBox("Infome o segundo n�mero: ")

On error resume next
c = a / b

if err.number <> 0 then
	MsgBox "Houve um erro, tente novamente"
else
	MsgBox "O valor da divis�o dos dois n�meros � " & c
end if

on error goto 0