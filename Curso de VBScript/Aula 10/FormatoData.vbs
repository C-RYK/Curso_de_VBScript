' FormatDateTime possui dois parâmetros:
'	Valor de tempo
'	Formato de saida para texto
'		0 - vbGeneralDate - Formato (dd/mm/aaaa hh:mi:ss) "Valor padrão"
'		1 - vbLongDate - Formato por extenso
'		2 - vbShortDate - Formato (dd/mm/aaaa)
'		3 - vbLongTime - Formato (hh:mi:ss)
'		4 - vbShortTime - Formato (hh:mi)

a = now()

'sgbox FormatDateTime(a)
msgbox FormatDateTime(a, 0)
msgbox FormatDateTime(a, 1)
msgbox FormatDateTime(a, 2)
msgbox FormatDateTime(a, 3)
msgbox FormatDateTime(a, 4)
