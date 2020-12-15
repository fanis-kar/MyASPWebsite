<!DOCTYPE html>
<html>

<head>
</head>

<body>
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.open = Application("connectionString")
Set cmd = CreateObject("ADODB.Command")

' Response.Write Now()
' Response.End

With cmd
.ActiveConnection = conn
.CommandType = 4 ' adCmdStoredProc
.CommandText = "MultipleInsert" ' Stored Procedure name

' cmd.Parameters("@FacultyAbbreviation") = "TestFacultyAbbreviation"
' cmd.Parameters("@FacultyName") = "TestFacultyName"
' cmd.Parameters("@FacultyWebsite") = "TestFacultyWebsite"
' cmd.Parameters("@FacultyEmail") = "TestFacultyEmail"
' cmd.Parameters("@FacultyPhone") = "TestFacultyPhone"
' cmd.Parameters("@UniversityId") = 1

' cmd.Parameters("@DepartmentAbbreviation") = "TestDepartmentAbbreviation"
' cmd.Parameters("@DepartmentName") = "TestDepartmentName"
' cmd.Parameters("@DepartmentWebsite") = "TestDepartmentWebsite"
' cmd.Parameters("@DepartmentEmail") = "TestDepartmentEmail"
' cmd.Parameters("@DepartmentPhone") = "TestDepartmentPhone"
' cmd.Parameters("@DepartmentCreated") = Now()

' .Parameters.Append cmd.CreateParameter("@FacultyName", adVarChar,adParamInput,255, "TestFacultyAbbreviation")

.Parameters.Append cmd.CreateParameter("@FacultyAbbreviation", 200,1,255, "TestFacultyAbbreviation") ' varchar
.Parameters.Append cmd.CreateParameter("@FacultyName", 200,1,255, "TestFacultyName") ' varchar
.Parameters.Append cmd.CreateParameter("@FacultyWebsite", 200,1,255, "TestFacultyWebsite") ' varchar
.Parameters.Append cmd.CreateParameter("@FacultyEmail", 200,1,255, "TestFacultyEmail") ' varchar
.Parameters.Append cmd.CreateParameter("@FacultyPhone", 200,1,255, "TestFacultyPhone") ' varchar
.Parameters.Append cmd.CreateParameter("@UniversityId", 3,1,100000, 1) ' integer
.Parameters.Append cmd.CreateParameter("@FacultyId", 3,2) ' outut

.Parameters.Append cmd.CreateParameter("DepartmentAbbreviation", 200,1,255, "TestDepartmentAbbreviation") ' varchar
.Parameters.Append cmd.CreateParameter("@DepartmentName", 200,1,255, "TestDepartmentName") ' varchar
.Parameters.Append cmd.CreateParameter("@DepartmentWebsite", 200,1,255, "TestDepartmentWebsite") ' varchar
.Parameters.Append cmd.CreateParameter("@DepartmentEmail", 200,1,255, "TestDepartmentEmail") ' varchar
.Parameters.Append cmd.CreateParameter("@DepartmentPhone", 200,1,255, "TestDepartmentPhone") ' varchar
.Parameters.Append cmd.CreateParameter("@DepartmentCreated", 135,1,8, Now()) ' datetime
.Parameters.Append cmd.CreateParameter("@DepartmentId", 3,2) ' outut

.Execute

FacultyId = cmd.Parameters("@FacultyId").Value
DepartmentId = cmd.Parameters("@DepartmentId").Value
End with

Response.Write "Stored Procedure Results <br>"
Response.Write "FacultyId: " & FacultyId & "<br>"
Response.Write "DepartmentId: " & DepartmentId & "<br>"

if err.number = 0 then
    Response.Write "ok"
else
    Response.Write err.Description
end if

Set cmd = Nothing
conn.close
%>
</body>
</html>