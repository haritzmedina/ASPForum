<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <%
    response.write("Logueatuta zaude " & request.cookies("erabizen") & ".<br>")
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
    sql = "select permisos from usuarios where nombre = '"&request.cookies("erabizen")&"';"
    set record = konexioa.execute(sql)

    response.write("Erabiltzaile maila: ")

    select case (record("permisos"))
      case 0:
        response.write("<font color=white>baimen gabekoa.</font>")
      case 1:
        response.write("<font color=blue>normala.</font>")
      case 2:
        response.write("<font color=#00FF00>moderatzailea.</font>")
      case 3:
        response.write("<font color=yellow>administratzailea.</font><br><a href='admin.asp?mod=0' target='body'>Administratzaile panela</a><br>")
      case else:
        response.write("<font class='error'>ez da egokia<br/>Administradorearekin hitz egin.")
    end select
      
    konexioa.close
    set konexioa = nothing
    set record =nothing
  %>
  <a href="perfila.asp" target="body">Perfila</a><br><br>
  <a href="logout.asp" target="_top">Irten</a>
</body>
</html>
