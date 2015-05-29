<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <%
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))

    sql = "select temaID from temas where temaid=(select max(temaid) from temas)"
    set record = konexioa.execute(sql)

    if record.eof <> 0 then
      temaid = 0
    else
      temaid = record("temaID")
    end if

    response.write("<a href=tema.asp?temaid="& temaid+1 & "&secid=" & request.querystring("id") & ">Tema berri bat sortu</a>")

    sql = "select * from temas where SeccionID = " & request.querystring("id")
    set record = konexioa.execute(sql)

    response.write("<table>")
    while (not record.eof)
      response.write("<tr><td><a href=tema.asp?temaid=" & record("TemaID") & "&secid=" & request.querystring("id") & ">" & record("Titulo") & "</a></td><td>" & record("fecha") & "</td></tr>")
      record.movenext
    wend

    response.write ("</table>")

    konexioa.close
    set konexioa = nothing
    set record = nothing

  %>
<br/>
<a href="menu.asp">Menu orokorrera bueltatu</a>
</body>
</html>
