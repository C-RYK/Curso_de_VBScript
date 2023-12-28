'INPUTBOX  - Pode ter até 7 parametros
'PROMPT  -> Mensagem que aparece no centro da caixa.
'TITLE -> Mensagem que aparece na barra de titulo da caixa.
'DEFAULT -> E um valor pré-definido da caixa.
'XPOS -> Posição X de coordenada carteziana em que a caixa aparecerá.
'YPOS -> Posição Y de coordenada carteziana em que a caixa aparecerá.
'HELPFILE -> É o arquivo de ajuda.
'CONTEXT -> É o contexto do arquivo de ajuda.


nome = inputBox("Diga seu nome", "Registro de pessoas", "Nome", 5555, 5555)
msgbox nome