<html>
<head>
    <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <h1>Erabiltzaileen login-a</h1>
  <form method="post" action="login.asp">
    Izena:<input type="text" name="izena"><br>
    Pasahitza:<input type="password" name="pass"><br>
    <input type="submit">
  </form>

  <%
    if request.form("izena") <> "" and request.form("pass") <> "" then
    dim konexioa, sql, record
    exist = 0 'Jakiteko erabiltzaileren bat dagoen datu basean
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
    sql = "select Nombre, Password from Usuarios"
    set record = konexioa.execute(sql)
    while (not record.eof)
      exist = 1
      if record("Nombre") = request.form("izena") and record("Password") = request.form("pass") then
        response.cookies("erabizen") = request.form("izena")
        response.cookies("erabpass") = request.form("pass")
        fecha = day(date) & "/" & month(date) & "/" & year(date) & " " & time
        sql = "update usuarios set ultimo_acceso = '"&fecha&"' where nombre='"&request.form("izena")&"';"
        set record = konexioa.execute(sql)
        response.redirect("frames.asp")
      else
        record.movenext
        if record.eof then
          response.write("Erabiltzailea ez da existitzen, edo pasahitza txarto<br>")
        end if
      end if
    wend
    if exist = 0 then
      response.write("Erabiltzaile bat sortu behar da lehenago<br>")
    end if
   
    konexioa.close
    set konexioa = nothing
    set record =nothing
    end if
  %>
<a href="berria.asp">Erabiltzaile berria egin</a><br/>
<a href="ezabatu.asp">Erabiltzailea ezabatu</a>
</body>
</html>
