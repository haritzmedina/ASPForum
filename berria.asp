<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <h1>Erabiltzaile berria</h1>
  <form method="post" action="berria.asp">
    Izena:&nbsp;<input type="text" name="izena"><br/>
    Pasahitza sartu:&nbsp;<input type="password" name="pass1"><br/>
    Berriro sartu:&nbsp;<input type="password" name="pass2"><br/>
    Jaiotze data:&nbsp;<select size=1 name="eguna">
    <%
      for a = 1 to 31
        response.write("<option>" & a & "</option>")
      next
    %>
    </select>&nbsp;
    <select size=1 name="hila">
    <%
      for a = 1 to 12
        response.write("<option>" & a & "</option>")
      next
    %>
    </select>&nbsp;
    <select size=1 name="urte">
      <%
        for a = year(date)-130 to year(date)
          response.write("<option>" & a & "</option>")
        next
      %>
    </select><br/>
    <input type="submit"></form>
  <%
    izena = request.form("izena")
    pass = request.form("pass1")
    data = request.form("eguna") & "/" & request.form("hila") & "/" & request.form("urte")
    if izena <> "" then
      if request.form("pass1") <> "" then
        if request.form("pass1") = request.form("pass2") then
          if request.form("urte") <> "" then
            dim konexioa, sql, record
            set konexioa = server.createobject("ADODB.connection")
            konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
            sql = "select * from Usuarios"
            set record = konexioa.execute(sql)
            if record.EOF <> 0 then
              sql = "insert into Usuarios (Nombre, password, nacimiento, permisos) values ('"&izena &"', '"&pass&"', '"&data&"', 3);"
            else
              sql = "insert into Usuarios (Nombre,password,nacimiento) values ('"&izena &"', '"&pass&"', '"&data&"');"
            end if
            set record = konexioa.execute(sql)
            konexioa.close
            set konexioa = nothing
            set record =nothing
            response.redirect("ondo.asp?id=1")
          else
            response.write("Urtea sartu egin behar da")
          end if
        else
          response.write("Pasahitzak ez dute koinziditzen")
        end if
      else
        response.write("Pasahitzak ezin dira utzik egon")
      end if
    end if
  %>
<a href="index.html">Login menura bueltatu.</a>
</body>
</html>
