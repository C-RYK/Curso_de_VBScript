erros = ""

a = InputBox("Informe o primeiro número: ")
b = InputBox("Informe o segundo número: ")

ValidaParametro(a)
ValidaParametro(b)

if len(trim(erros)) > 0 then
	msgbox erros, 16, "ATENÇÃO"

else
	msgbox "A soma do primeiro com o segundo número é " & (cdbl(a) + cdbl(a)), 64, "Resultado"
end if

' -----------------------------------------------------------------------------------------------------------------------------------------

sub validaParametro(Parametro)
	if not Isnumeric(Parametro) then
		erros = erros & "Os valores devem ser númericos" & chr(10)
	else
		if Parametro < 0 then
			erros = erros & "Os valores devem ser maiores que zero" & chr(10)
		end if
	end if
end sub
