<html>
<head>
</head>
<body>
  <%
    response.cookies("erabizen") = ""
    response.cookies("erabpass") = ""
    response.redirect("index.html")
  %>
</body>
</html>
