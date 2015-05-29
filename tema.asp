<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <%
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))

    sql = "select temaID from temas where temaID=" & request.querystring("temaid")
    set record = konexioa.execute(sql)

    temaid = request.querystring("temaid")
    secid = request.querystring("secid")

    if record.Eof <> 0 then
      %>
      <form method="post" action="postit.asp?new=1&temaid=<%=temaid%>&secid=<%=secid%>"><!-- Ondo funtzionatzen du, nahiz eta kodigoan ondo ez ikusi -->
        Temaren gaia: <input name="titulo" type="text" size="75"><br/>
        Textua: <textarea name="post" rows=10 cols=40></textarea><br/>
        <input type="submit">
      </form>
      <%
    else
      sql = "select * from posts where temaID=" & request.querystring("temaid")
      set record = konexioa.execute(sql)
      %>
      <table>
      <%
      while (not record.eof)
        sql2 = "select nombre from usuarios where ID=" & record("Usuario")
        set record2 = konexioa.execute(sql2)
        response.write("<tr rows=2><td>" & record2("nombre") & "</td><td>" & record("Fecha") & "</td></tr><tr><td>" & record("Asunto") & "</td></tr><tr><td>" & record("Post") & "</td></tr>")
        record.movenext
      wend

      konexioa.close
      set konexioa = nothing
      set record = nothing
      set record2 = nothing

      %>

      <tr><td><form method="post" action="postit.asp?new=0&temaid=<%=temaid%>&secid=<%=secid%>"><!-- Ondo funtzionatzen du, nahiz eta kodigoan ondo ez ikusi -->
        Temaren gaia: </td><td><input name="titulo" type="text" size="75"><br/></td></tr>
        <tr><td>Textua: </td><td><textarea name="post" rows=10 cols=40></textarea><br/>
        <input type="submit">
      </form></td></tr></table>

      <%
    end if
  %>
  <a href="seccion.asp?id=<%=secid%>">Atzera bueltatu</a>
</body>
</html>
