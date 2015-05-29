<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <h1>Erabiltzailea ezabatu</h1>
  <form method="post" action="ezabatu.asp">
    Izena:<input type="text" name="izena"><br/>
    Pasahitza:<input type="password" name="pass"><br/>
    Ziur zaude?:<input type="checkbox" name="ezabatu">
    <input type="submit">
  </form>

  <%
    if request.form("izena") <> "" and request.form("pass") <> "" then
    response.write(request.form("ezabatu"))
      if request.form("ezabatu") = "on" then
        dim konexioa, sql, record
        set konexioa = server.createobject("ADODB.connection")
        konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
        sql = "delete from Usuarios where Nombre='"&request.form("izena")&"' and Password='"&request.form("pass")&"';"
        set record = konexioa.execute(sql)
        konexioa.close
        set konexioa = nothing
        set record =nothing
        response.redirect("login.asp")
      else
        response.write("Ez duzu ziurtasunik eman borratzeko.")
      end if
    end if

  %>
  
</body>
</html>
