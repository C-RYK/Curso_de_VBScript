Set tb = createObject("ADODB.Connection")
Set rc = createObject("ADODB.RecordSet")

tb.opne "Provider=Microsoft.jet.OLEDB.4.0; Data source=C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Banco de Dados\DBVBScript.mdb"

VCodInfPessoa = ""
VNome = ""
VDDD = ""
VTelefone = ""
VTipoTelefone = ""
erros = ""
VResultado = ""

Do While true
	
	wscript.stdout.write "Informe o código da pessoa [" & VCodInfPessoa & "]"
	VCodInfPessoa = wscript.stdin.readline

	wscript.stdout.write "Informe o nome [" & VNome & "]"
	VNome = wscript.stdin.readline

	wscript.stdout.write "Informe o DDD [" & VDDD & "]"
	VDDD = wscript.stdin.readline

	wscript.stdout.write "Informe o Telefone [" & VTelefone & "]"
	VTelefone = wscript.stdin.readline

	wscript.stdout.write "Informe o tipo de Telefone [" & VTipoTelefone & "]"
	VTipoTelefone = wscript.stdin.readline

	validaCampos

	if len(trim(erros)) = 0 then
		tb.execute "insert into agenda (CodInfPessoa, " & _
			        "             Nome, " & _
			        "             DDD, " & _
			        "             Telefone, " & _
			        "             TipoTelefone) " & _
			        " Value (" & VCodInfPessoa &", " & _
			        "             " & VNome & " , " " & _
                                      "             " & VDDD & ", " " & _
	                            "             " & VTelefone & ", " & _
                                      "             " & VTipoTelefone & " )"
		Exit do
	else
		Resposta = msgbox("Ocorreram erros no processamento: " & chr(10) & erros, 16 + 5, "ATENÇÃO")
		erros = ""
		
		if Resposta <> 4 then
			exit do
		end if

	end if
Loop

rc.CursorLocation = 3

rc.Open "Select = from agenda", tb, 2,3

rc.MoveFirst

Do while not rc.eof
	VResultado = VResultado & "Código - [" & rc.fileds("codinfpessoa").value & "] " & _
					         "Código - [" & rc.fileds("nome").value & "] " & _
                                                           "Código - [" & rc.fileds("ddd").value & "] " & _
                                                           "Código - [" & rc.fileds("telefone").value & "] " & _
                                                           "Código - [" & rc.fileds("tipotelefone").value & "] " &  chr(10)

	rc.MoveNext
Loop

msgbox VResultado

sub validaCampos
	'CodInfPessoa - Numérico, inteiro, >0
	if not Isnumeric(VCodInfPessoa) then
		erros = erros & "O campo código deve ser númerico" & chr(10)
	else
		if cdbl(VCodInfPessoa) <= 0 then
			erros = erros & "O campo código deve ser maior que zero" & chr(10)
		end if

		if cdbl(VCodInfPessoa) <> round(cdbl(VCodInfPessoa)) then
			erros = erros & "O campo código deve ser inteiro" & chr(10)
		end if
	end if

	'DDD - Numérico, inteiro, >0, entre 10 e 99
		if not Isnumeric(VDDD) then
			erros = erros & "O campo DDD deve ser númerico" & chr(10)
		else
			if cdbl(VDDD) <= 0 then
				erros = erros & "O campo DDD deve ser maior que zero" & chr(10)
			end if

			if cdbl(VDDD) <> round(cdbl(VDDD)) then
				erros = erros & "O campo código deve ser inteiro" & chr(10)
			end if
		
			if not (cdbl(VDDD)) >= 10 and cdbl(VDDD) <= 99) then
				erros = erros & "O campo DDD deve estar entre 10 e 99" chr(10)
			end if
		end if

	'Telefone - Numérico, inteiro, >0, entre 100000000 e 999999999
		if not Isnumeric(VTelefone) then
			erros = erros & "O campo Telefone deve ser númerico" & chr(10)
		else
			if cdbl(VTelefone) <= 0 then
				erros = erros & "O campo Telefone deve ser maior que zero" & chr(10)
			end if

			if cdbl(VTelefone) <> round(cdbl(VTelefone)) then
				erros = erros & "O campo Telefone deve ser inteiro" & chr(10)
			end if

			if not (cdbl(VTelefone) >= 100000000 and cdbl(VTelefone) <= 999999999) then
				erros = erros & "O campo Telefone deve estar entre 100000000 e 900000000" & chr(10)
			end if

		end if

	'Nome
	if len(trim(VNome)) = 0 then
		erros = erros & "O campo Nome é obrigatorio" & chr(10)
	end if

	'Tipo
	if len(trim(VTipoTelefone)) = 0 then
		erros = erros & "O campo tipo é obrigadorio" & chr(10)
	end if

end sub