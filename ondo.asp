<html>
<head>
  <link rel="stylesheet" type="text/css" href="styles/estiloa.css">
</head>
<body>
  <table>
    <tr><td><font class="goi">Mezua</font></td></tr>
    <tr><td><font class="ondomezua">
    <%
      if request.querystring("id") = 1 then
        response.write("Ondo erregistratuta dago.<br/><a href='index.html'>Egin klik hemen bueltatzeko.</a>")
      elseif request.querystring("id") = 2 then
        response.write("Ondo borratu da sekzioa.<br/><a href='admin.asp'>Egin klik hemen bueltatzeko.</a>")
      end if
    %>
    </font></td></tr>
  </table>
</body>
</html>
