' IsNumeric - Possui um par�metro que � o valor  a ser analisado
'	Validar se o valor informado � um n�mero

' IsDate - Possui um par�metro que � o valor  a ser analisado
'	Validar se o valor informado � uma data

' IsNull - Possui um par�metro que � o valor  a ser analisado
'	Validar se o valor informado � um nulo

' IsEmpty - Possui um par�metro que � o valor  a ser analisado
'	Validar se a variavel est� vazia

a = Inputbox("D�gite um valor qualquer: ")
b = null

if Isnumeric(a) then
	msgbox "O valor informado � n�merico"

elseif IsDate(a) then
	msgbox "O valor informado � uma data"

else
	msgbox "O valor informado � caractere"
End if

if Isnull(b) then
	msgbox "O valor B � igual a Nulo"
end if

if IsEmpty(c) then
	msgbox "O valor C � vazio"
end if
