Set eng = createobject("DAO.DBEngine.36")

Set tb = eng.createDataBase("C:\Users\carlo\OneDrive\Documentos\Curso de VBScript\Banco de Dados\DBVBScript.mdb", _
				        "jet 3. x; LANGID=0x409; CP=1252; COUNTRY=0")

tb.execute "create table agenda (codInfPessoa Integer, " & _
					  "Nome String, " & _
					  "DDD Integer, " & _
				            "Telefone Integer, " & _
					  "TipoTelefone String)"

tb.execute "insert into agenda (codInfPessoa, " & _
	        "			           Nome, " & _
	        "				 DDD, " & _
	        "				 Telefone, " & _
	        "				 TipoTelefone)" & _
	        "		Values (1, " & _
	        "			    Carlos Henrique, " & _
	        "			     11, " & _
	        "			      954544761" & _
	        "			      ' Celular ')"

Set rc = tb.OpenRecordSet("Select * from agenda")

rc.MoveFirst

msgbox "Código - ["& rc.Fields("codInfPessoa").value &"]" & chr(10) & _
	    "Nome - ["& rc.Fields("Nome").value &"]" & chr(10) & _
	    "DDD - ["& rc.Fields("DDD").value &"]" & chr(10) & _
              "Telefone - ["& rc.Fields("Telefone").value &"]" & chr(10) & _
	    "Tipo - ["& rc.Fields("TipoTelefone").value &"]"

rc.close
tb.close
