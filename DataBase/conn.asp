<%
Server.ScriptTimeout = "10"
set conn=server.CreateObject("Adodb.Connection")
Path="driver={SQL Server};server=(local);uid=sa;pwd=123456;database=SchoolManage"
conn.open Path

%>