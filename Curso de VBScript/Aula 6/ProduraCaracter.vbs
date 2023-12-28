'instr possui até 3 parâmetros
'	Caracter inicial (Opcional, caso não seja informado assume "1")
'	Texto original
'	Texto procurado

z = inputbox("Informe o indice do caracter inicial: ") + 0
a = inputBox("Informe qualquer valor: ")
b = inputBox("Caracter de procura: ")

c = instr(z, a, b)

msgbox "O(s) caracter(es) " & b & "está na posição " & c