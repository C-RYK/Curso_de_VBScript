' FormatNumber - possui at� 5 par�metros
'	Expression - que � o n�mero a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou n�o para os casos de decimais
'		-2 - O valor ser� o configurado  no windows
'		-1 - O valor zero aparecer� em casos de fra��es menores que 1 inteiro
'		0 - O valor zero n�o aparecer� em casos de fra��es menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecer�o com sinal de "-" ou entre "()"
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta n�mero entre par�nteses
'		0 - Apresenta o n�mero com o sinal de "-"
'	GroupDig - Indica se aparecer� o ponto de separa��o de milhar, milh�o etc.
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta o ponto
'		0 - N�o apresenta o n�mero

a = 0.5

msgbox FormatNumber(a)
msgbox FormatNumber(a, 3, 0)

a = 20530.53

msgbox FormatNumber(a, 2, , ,-1)
msgbox FormatNumber(a, 3, , ,0)

a = -1020530.53

msgbox FormatNumber(a, 2, ,-1,-1)
msgbox FormatNumber(a, 3, ,0,0)




' FormatCurrency - possui at� 5 par�metros
'	Expression - que � o n�mero a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou n�o para os casos de decimais
'		-2 - O valor ser� o configurado  no windows
'		-1 - O valor zero aparecer� em casos de fra��es menores que 1 inteiro
'		0 - O valor zero n�o aparecer� em casos de fra��es menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecer�o com sinal de "-" ou entre "()"
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta n�mero entre par�nteses
'		0 - Apresenta o n�mero com o sinal de "-"
'	GroupDig - Indica se aparecer� o ponto de separa��o de milhar, milh�o etc.
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta o ponto
'		0 - N�o apresenta o n�mero

a = 0.5

msgbox FormatCurrency(a)
msgbox FormatCurrency(a, 3, 0)

a = 20530.53

msgbox FormatCurrency(a, 2, , ,-1)
msgbox FormatCurrency(a, 3, , ,0)

a = -1020530.53

msgbox FormatCurrency(a, 2, ,-1,-1)
msgbox FormatCurrency(a, 3, ,0,0)




' FormatPercent - possui at� 5 par�metros
'	Expression - que � o n�mero a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou n�o para os casos de decimais
'		-2 - O valor ser� o configurado  no windows
'		-1 - O valor zero aparecer� em casos de fra��es menores que 1 inteiro
'		0 - O valor zero n�o aparecer� em casos de fra��es menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecer�o com sinal de "-" ou entre "()"
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta n�mero entre par�nteses
'		0 - Apresenta o n�mero com o sinal de "-"
'	GroupDig - Indica se aparecer� o ponto de separa��o de milhar, milh�o etc.
'		-2 - O valor ser� o configurado no windows
'		-1 - Apresenta o ponto
'		0 - N�o apresenta o n�mero

a = 1

msgbox FormatPercent(a)
msgbox FormatPercent(a, 3, 0)

a = 0.5

msgbox FormatPercent(a, 2, , ,-1)
msgbox FormatPercent(a, 3, , ,0)

a = 0.05

msgbox FormatPercent(a, 2, ,-1,-1)
msgbox FormatPercent(a, 3, ,0,0)
