erros = ""

a = InputBox("Informe o primeiro n�mero: ")
b = InputBox("Informe o segundo n�mero: ")

if not validaParametro(a, b) then
	msgbox erros, 16, "ATEN��O"
else
	msgbox "A soma do primeiro com o segundo n�mero � " & (cdbl(a) + cdbl(b)), 64, "Resultado"
end if

' ------------------------------------------------------------------------------------------------------------------------------------

function validaParametro(Parametro1, Parametro2)
	If not Isnumeric(parametro1) then
		erros = erros & "O valor do parametro 1 deve ser n�merico" & chr(10)
	else
		if parametro1 < 0 then
			erros = erros & "O valor do par�metro 1 deve ser maior que zero" & chr(10)
		end if
	end if

	if not Isnumeric(Paremetro2) then
		erros = erros & "O valor do par�metro 2 deve ser n�merico" & chr(10)
	else
		if Parametro2 < 0 then
			erros = erros & "O valor do par�metro 2 deve ser maior que zero" & chr(10)
		end if
	end if
	
	if len(trim(erros)) > 0 then
		validaParametro = false
	else
		validaParametro = true
	end if
end function
