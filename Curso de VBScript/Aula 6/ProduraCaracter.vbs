'instr possui at� 3 par�metros
'	Caracter inicial (Opcional, caso n�o seja informado assume "1")
'	Texto original
'	Texto procurado

z = inputbox("Informe o indice do caracter inicial: ") + 0
a = inputBox("Informe qualquer valor: ")
b = inputBox("Caracter de procura: ")

c = instr(z, a, b)

msgbox "O(s) caracter(es) " & b & "est� na posi��o " & c