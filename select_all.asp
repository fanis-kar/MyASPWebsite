<!DOCTYPE html>
<html>

<head>
</head>

<body>

<!--#include file="include_header.asp"-->
<%
set conn = Server.CreateObject("ADODB.Connection")
conn.open = Application("connectionString") 
set rs = CreateObject("ADODB.Recordset")

query= "SELECT * FROM Universities"

rs.Open query, conn

do until rs.EOF
  for each row in rs.Fields
    Response.Write(row.name)
    Response.Write(" = ")
    Response.Write(row.value & "<br>")
  next
  Response.Write("<br>")
  rs.MoveNext
loop

rs.close
conn.close
%>
</body>
</html>