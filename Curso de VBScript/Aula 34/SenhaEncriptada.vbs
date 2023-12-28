Set oEncoder = createObject("Scripting.Encoder")
senhaEncriptada = oEncoder.EncodeScriptFile(".vbs", InputBox("Informe sua senha: "), 0, "")

MsgBox senhaEncriptada
