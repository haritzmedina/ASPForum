<html>
<head>
    <link rel="stylesheet" type="text/css" href="styles/estiloadmin.css">
</head>
<body>
  <%
    dim konexioa, sql, record
    set konexioa = server.createobject("ADODB.connection")
    konexioa.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("foro.mdb"))
    sql = "select permisos from Usuarios where nombre = '" & request.cookies("erabizen") & "' and password='"&request.cookies("erabpass")&"';"
    set record = konexioa.execute(sql)

    if record("permisos") = 3 then
      if request.querystring("mod") = 0 then 'Menu orokorra
        %>
          <font class="advert">KONTUZ HEMEN EGITEN DUZUNAREKIN</font>
          <table>
            <tr><td><a href="admin.asp?mod=1">Erabiltzaileak kudeatu</a></td></tr>
            <tr><td><a href="admin.asp?mod=2">Foroaren sekzioak kudeatu</a></td></tr>
            <tr><td><a href="menu.asp">Menura bueltatu</a></td></tr>
          </table>
        <%
      elseif request.querystring("mod") = 1 then 'Erabiltzaileen kudeaketa
        sql2 = "select * from Usuarios"
        set record2 = konexioa.execute(sql2)
        response.write("<div align='center'><font class='title'>Erabiltzaileen kudeaketa</font><br/><br/><table><tr><td><font class='top'>Erabiltzailea</font></td><td><font class='top'>Aldatu</font></td><td><font class='top'>Ezabatu</font></td>")
        while (not record2.eof)
          response.write("<tr><td class='users'><font class=izena>" & record2("nombre") & "</font></td><td class='users'><a href=admin.asp?mod=4&id="&record2("id")&"><img src='images/edit.png'></a></td><td class='users'>")
          if record2("permisos") <> 3 then
            response.write("<a href=admin.asp?mod=3&id="&record2("id") & "><img src='images/delete.png'></a>")
          else
            response.write("&nbsp;")
          end if
          response.write("</td></tr>")
          record2.movenext
        wend
        response.write("</table></div>")
        set record2 = nothing
      elseif request.querystring("mod") = 2 then 'Sekzio kudeaketa
        %>
          <form action="admin.asp?mod=2" method="post">
            <font color=white>Sekzioaren izena:</font> <input type="text" name="sec"><br/>
            <Font color=white>Deskripzioa:</font> <textarea name="des" rows=5 cols=25></textarea><br/>
            <input type="submit" name="sub" value="Sortu">
            <input type="submit" name="sub" value="Borratu">
          </form>
        <%
        if request.form <> "" then
          if request.form("sub") = "Sortu" then
            sql2 = "insert into secciones (Nombre, descripcion) values ('"&request.form("sec")&"', '"&request.form("des")&"');"
            set record2 = konexioa.execute(sql2)
            response.redirect("admin.asp")
          elseif request.form("sub") = "Borratu" then
            sql2 = "select * from secciones where nombre = '"&request.form("sec")&"';"
            set record2 = konexioa.execute(sql2)
            if (not record2.eof) then
              sql3 = "delete from secciones where nombre='"&request.form("sec")&"';"
              set record3 = konexioa.execute(sql3)
              response.redirect("ondo.asp?id=2")
            else
              response.write("<font color=white>Sekzioa ez da existitzen.</font>")
            end if
          else
            response.write("Programazio errorea")
          end if
        end if
      elseif request.querystring("mod") = 3 then 'Erabiltzailea borratu
        sql2 = "delete from Usuarios where id=" &request.querystring("id")&";"
        set record2 = konexioa.execute(sql2)
        sql2 = "update posts set asunto='Ezer', post='Erabiltzailea kendu egin du', usuario=0 where usuario=" &request.querystring("id")&";"
        set record2 = nothing
      elseif request.querystring("mod") = 4 then 'Erabiltzailea editatu
        sql2 = "select * from usuarios where id=" & request.querystring("id")&";"
        set record2 = konexioa.execute(sql2)
        %>
          <form action='admin.asp?mod=4&id=<%=request.querystring("id")%>' method="post">
            <font color=#FFFFFF>Izena: </font><input type="text" name="izena" value='<%=record2("nombre")%>'><br/>
            <font color=#FFFFFF>Baimenak: </font><input type="text" name="baimen" value='<%=record2("permisos")%>'><br/>
            <input type="submit" value="Eguneratu">
          </form>
        <%
          if request.form() <> "" then
            sql3 = "update usuarios set nombre='"&request.form("izena")&"', permisos="&request.form("baimen")&" where id="&request.querystring("id")&";"
            set record3 = konexioa.execute(sql3)
            response.redirect("admin.asp")
          end if
      end if
    else %>
    <font class="error">Ez dituzu baimenik sartzeko</font><br/><a href="index.html" target="_top">Hasierara bueltatu</a>
    <%
    end if
  %>
</body>
</html>
