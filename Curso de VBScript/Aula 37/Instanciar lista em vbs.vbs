dim list

Set list = CreateObject("System.Collections.ArrayList")

list.Add "Carlos"

list.Add "Henrique"

list.Add "Parreira"

list.Add "Jacinto"

list.Reverse

wscript.echo list.Count                 ' --> 4

wscript.echo list.Item(0)               ' --> Jacinto

wscript.echo list.IndexOf("Parreira", 0)   ' --> 1

wscript.echo join(list.ToArray(), ", ") ' --> Jacinto, Parreira, Henrique, Carlos


