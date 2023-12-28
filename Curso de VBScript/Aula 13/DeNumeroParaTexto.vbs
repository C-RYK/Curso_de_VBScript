' FormatNumber - possui até 5 parâmetros
'	Expression - que é o número a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou não para os casos de decimais
'		-2 - O valor será o configurado  no windows
'		-1 - O valor zero aparecerá em casos de frações menores que 1 inteiro
'		0 - O valor zero não aparecerá em casos de frações menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecerão com sinal de "-" ou entre "()"
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta número entre parênteses
'		0 - Apresenta o número com o sinal de "-"
'	GroupDig - Indica se aparecerá o ponto de separação de milhar, milhão etc.
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta o ponto
'		0 - Não apresenta o número

a = 0.5

msgbox FormatNumber(a)
msgbox FormatNumber(a, 3, 0)

a = 20530.53

msgbox FormatNumber(a, 2, , ,-1)
msgbox FormatNumber(a, 3, , ,0)

a = -1020530.53

msgbox FormatNumber(a, 2, ,-1,-1)
msgbox FormatNumber(a, 3, ,0,0)




' FormatCurrency - possui até 5 parâmetros
'	Expression - que é o número a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou não para os casos de decimais
'		-2 - O valor será o configurado  no windows
'		-1 - O valor zero aparecerá em casos de frações menores que 1 inteiro
'		0 - O valor zero não aparecerá em casos de frações menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecerão com sinal de "-" ou entre "()"
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta número entre parênteses
'		0 - Apresenta o número com o sinal de "-"
'	GroupDig - Indica se aparecerá o ponto de separação de milhar, milhão etc.
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta o ponto
'		0 - Não apresenta o número

a = 0.5

msgbox FormatCurrency(a)
msgbox FormatCurrency(a, 3, 0)

a = 20530.53

msgbox FormatCurrency(a, 2, , ,-1)
msgbox FormatCurrency(a, 3, , ,0)

a = -1020530.53

msgbox FormatCurrency(a, 2, ,-1,-1)
msgbox FormatCurrency(a, 3, ,0,0)




' FormatPercent - possui até 5 parâmetros
'	Expression - que é o número a ser transformado
'	NumDigAfterDec - Quantidade de decimais.
'	IncLeadingDig - Indica se o zero aparece ou não para os casos de decimais
'		-2 - O valor será o configurado  no windows
'		-1 - O valor zero aparecerá em casos de frações menores que 1 inteiro
'		0 - O valor zero não aparecerá em casos de frações menores que 1 inteiro
'	UseParForNegNum - Indica se valores negativos aparecerão com sinal de "-" ou entre "()"
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta número entre parênteses
'		0 - Apresenta o número com o sinal de "-"
'	GroupDig - Indica se aparecerá o ponto de separação de milhar, milhão etc.
'		-2 - O valor será o configurado no windows
'		-1 - Apresenta o ponto
'		0 - Não apresenta o número

a = 1

msgbox FormatPercent(a)
msgbox FormatPercent(a, 3, 0)

a = 0.5

msgbox FormatPercent(a, 2, , ,-1)
msgbox FormatPercent(a, 3, , ,0)

a = 0.05

msgbox FormatPercent(a, 2, ,-1,-1)
msgbox FormatPercent(a, 3, ,0,0)
