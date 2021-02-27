<%
response.expires = -1

q = ucase(request.querystring("q"))

set conn = Server.CreateObject("ADODB.Connection")
conn.open = Application("connectionString") 
set rs = CreateObject("ADODB.Recordset")

sql = "SELECT * FROM Universities WHERE Abbreviation LIKE '%" & q & "%'"

rs.Open sql, conn, 3, 1

Response.ContentType = "application/json"

do while not rs.EOF
  Response.Write("{ ""Id"": " & rs("Id") & ", ""Abbreviation"": """ & rs("Abbreviation") & """, ""Name"": """ & rs("Name") & """, ""Address"": """ & rs("Address") & """ },")

  rs.MoveNext

  'exit do
loop


%>