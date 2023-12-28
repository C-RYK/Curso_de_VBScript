' IsNumeric - Possui um parâmetro que é o valor  a ser analisado
'	Validar se o valor informado é um número

' IsDate - Possui um parâmetro que é o valor  a ser analisado
'	Validar se o valor informado é uma data

' IsNull - Possui um parâmetro que é o valor  a ser analisado
'	Validar se o valor informado é um nulo

' IsEmpty - Possui um parâmetro que é o valor  a ser analisado
'	Validar se a variavel está vazia

a = Inputbox("Dígite um valor qualquer: ")
b = null

if Isnumeric(a) then
	msgbox "O valor informado é númerico"

elseif IsDate(a) then
	msgbox "O valor informado é uma data"

else
	msgbox "O valor informado é caractere"
End if

if Isnull(b) then
	msgbox "O valor B é igual a Nulo"
end if

if IsEmpty(c) then
	msgbox "O valor C é vazio"
end if
