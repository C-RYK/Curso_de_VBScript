erros = ""

a = InputBox("Informe o primeiro número: ")
b = InputBox("Informe o segundo número: ")

if not validaParametro(a, b) then
	msgbox erros, 16, "ATENÇÃO"
else
	msgbox "A soma do primeiro com o segundo número é " & (cdbl(a) + cdbl(b)), 64, "Resultado"
end if

' ------------------------------------------------------------------------------------------------------------------------------------

function validaParametro(Parametro1, Parametro2)
	If not Isnumeric(parametro1) then
		erros = erros & "O valor do parametro 1 deve ser númerico" & chr(10)
	else
		if parametro1 < 0 then
			erros = erros & "O valor do parâmetro 1 deve ser maior que zero" & chr(10)
		end if
	end if

	if not Isnumeric(Paremetro2) then
		erros = erros & "O valor do parâmetro 2 deve ser númerico" & chr(10)
	else
		if Parametro2 < 0 then
			erros = erros & "O valor do parâmetro 2 deve ser maior que zero" & chr(10)
		end if
	end if
	
	if len(trim(erros)) > 0 then
		validaParametro = false
	else
		validaParametro = true
	end if
end function
