erros = ""

a = InputBox("Informe o primeiro n�mero: ")
b = InputBox("Informe o segundo n�mero: ")

ValidaParametro(a)
ValidaParametro(b)

if len(trim(erros)) > 0 then
	msgbox erros, 16, "ATEN��O"

else
	msgbox "A soma do primeiro com o segundo n�mero � " & (cdbl(a) + cdbl(a)), 64, "Resultado"
end if

' -----------------------------------------------------------------------------------------------------------------------------------------

sub validaParametro(Parametro)
	if not Isnumeric(Parametro) then
		erros = erros & "Os valores devem ser n�mericos" & chr(10)
	else
		if Parametro < 0 then
			erros = erros & "Os valores devem ser maiores que zero" & chr(10)
		end if
	end if
end sub
