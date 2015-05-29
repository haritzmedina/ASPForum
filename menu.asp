<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <%
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
    sql = "select * from secciones"
    set record = konexioa.execute(sql)
    
    response.write("<table>")
    while (not record.eof)
      response.write("<tr><td><a href=seccion.asp?id="& record("SeccionID") &">" & record("Nombre") & "</a></tr></td><tr><td>" & record("Descripcion") & "</td></tr>")
      record.movenext
    wend

    response.write("</table>")
    
    konexioa.close
    set konexioa = nothing
    set record = nothing
  %>
</body>
</html>
