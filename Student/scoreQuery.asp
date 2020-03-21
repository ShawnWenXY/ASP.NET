<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/function.asp"-->
<%
  '如果学生还没有登陆
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>学生成绩查询</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="scoreQuery.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>成绩信息管理--&gt;学生成绩查询
			 </td>
	   </tr>
		 <tr>
		   <td>请选择学期:</td>
			 <td>
			   <select name=termId>
				    <option value="">请选择</option>
					  <%
					    dim sqlString
						  sqlString = "select * from [termInfo]"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>&nbsp;<input type="submit" name="submit" value="查询成绩">
				</td>
		 </tr>
		 <tr>
		   <td colspan=2>
			   <table width=100% border=0 bordercolor="maroon" cellspacing=0>
				  <%
				     if Request("termId") <> "" then
						    Response.Write "<tr><td colspan=8 align=center>" & GetTermNameById(Request("termId")) & " 成绩信息</td></tr>"
							  Response.Write "<tr><td>课程编号</td><td>课程名称</td><td>课程类型</td><td>成绩</td></tr>"
					 
						   '查询必修课的成绩
						   sqlString = "select [scoreInfo].courseNumber,[scoreInfo].score,[classCourseInfo].courseName from [scoreInfo],[classCourseInfo] where [scoreInfo].studentNumber='" & session("studentNumber") & "' and [scoreInfo].isSelect=0 and [scoreInfo].courseNumber=[classCourseInfo].courseNumber and [classCourseInfo].termId=" & Request("termId")
						   set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
						   scoreInfoRs.Open sqlString,conn,1,1
						   while not scoreInfoRs.EOF
						     Response.Write "<tr><td>" & scoreInfoRs("courseNumber") & "</td><td>" & scoreInfoRs("courseName") & "</td><td>必修课</td><td>" & scoreInfoRs("score") & "</td></tr>"
							   scoreInfoRs.MoveNext
						   wend
						   scoreInfoRs.Close
						   '查询选修课的成绩
						   sqlString = "select [scoreInfo].courseNumber,[scoreInfo].score,[publicCourseInfo].courseName from [scoreInfo],[publicCourseInfo] where [scoreInfo].studentNumber='" & session("studentNumber") & "' and [scoreInfo].isSelect=1 and [scoreInfo].courseNumber=[publicCourseInfo].courseNumber and [publicCourseInfo].termId=" & Request("termId")
						   scoreInfoRs.Open sqlString,conn,1,1
						   while not scoreInfoRs.EOF
						     Response.Write "<tr><td>" & scoreInfoRs("courseNumber") & "</td><td>" & scoreInfoRs("courseName") & "</td><td>选修课</td><td>" & scoreInfoRs("score") & "</td></tr>"
							   scoreInfoRs.MoveNext
						   wend
						   scoreInfoRs.Close
						end if
				  %>
				 </table>
		   </td>
		 </tr>
		</form>
	</table>
</body>
</html>
		