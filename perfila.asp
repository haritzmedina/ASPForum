<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <%
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
    sql = "select * from Usuarios where Nombre = '"&request.cookies("erabizen")&"';"
    set record = konexioa.execute(sql)

  %>

  <form method="post" action='perfila.asp'>
    Izena:&nbsp;<input type="text" name="izena" value='<%=record("nombre")%>'><br/>
    Pasahitza berria:&nbsp;<input type="password" name="pass1" value='<%=record("password")%>'><br/>
    Berriro sartu:&nbsp;<input type="password" name="pass2" value='<%=record("password")%>'><br/>
    <input type="submit"></form>
    <%
      if request.form() <> "" then
        if request.form("eguna") <> "" and request.form("hila") <> "" and request.form("urte") <> "" then
          data = request.form("eguna") & "/" & request.form("hila") & "/" & request.form("urte")
        else
           data = record("nacimiento")
        end if
        if request.form("pass1") = request.form("pass2") then
          sql = "update usuarios set izena='"&request.form("Izena")&"';"
          set record = konexioa.execute(sql)
        else
          response.write("Pasahitzak ez dute koinziditzen")
        end if
      end if
    %>
</body>
</html>
