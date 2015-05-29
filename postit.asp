<html>
<head>
</head>
<body>
<%

  dim konexioa, sql, record
  set konexioa = server.createobject("ADODB.connection")
  konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))

  if request.querystring("new") = 1 then
    sql = "insert into temas (titulo, seccionID) values ('"& request.form("titulo") &"', " & request.querystring("secid") & ");"
    set record = konexioa.execute(sql)
  end if

  sql = "select * from usuarios where nombre='"&request.cookies("erabizen")&"';"
  set record = konexioa.execute(sql)

  user = record("id")

  fecha = day(date) & "/" & month(date) & "/" & year(date) & " " & time

  sql = "insert into posts (Usuario, Asunto, Fecha, Post, TemaID) values ("&user&", '" & request.form("titulo") & "', '"& fecha &"', '" & request.form("post") & "', "& request.querystring("temaid") &");"
  set record = konexioa.execute(sql)

    konexioa.close
    set konexioa = nothing
    set record = nothing

    response.redirect("tema.asp?temaid="& request.querystring("temaid")&"&secid="&request.querystring("secid"))
%>
</body>
</html>
