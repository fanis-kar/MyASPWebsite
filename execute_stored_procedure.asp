<!DOCTYPE html>
<html>

<head>
</head>

<body>
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open = Application("connectionString") 
Set cmd = CreateObject("ADODB.Command")

With cmd
.ActiveConnection = conn
.CommandType = 4 ' adCmdStoredProc
.CommandText = "MultipleInsert" ' Stored Procedure name

cmd.Parameters("@FacultyAbbreviation") = "TestFacultyAbbreviation"
cmd.Parameters("@FacultyName") = "TestFacultyName"
cmd.Parameters("@FacultyWebsite") = "TestFacultyWebsite"
cmd.Parameters("@FacultyEmail") = "TestFacultyEmail"
cmd.Parameters("@FacultyPhone") = "TestFacultyPhone"
cmd.Parameters("@UniversityId") = 1

cmd.Parameters("@DepartmentAbbreviation") = "TestDepartmentAbbreviation"
cmd.Parameters("@DepartmentName") = "TestDepartmentName"
cmd.Parameters("@DepartmentWebsite") = "TestDepartmentWebsite"
cmd.Parameters("@DepartmentEmail") = "TestDepartmentEmail"
cmd.Parameters("@DepartmentPhone") = "TestDepartmentPhone"

.Execute

FacultyId = cmd.Parameters("@FacultyId").Value
DepartmentId = cmd.Parameters("@DepartmentId").Value
End with

Response.Write "Stored Procedure Results <br>"
Response.Write "FacultyId: " & FacultyId & "<br>"
Response.Write "DepartmentId: " & DepartmentId & "<br>"

Set cmd = Nothing
conn.close 
%>
</body>
</html>