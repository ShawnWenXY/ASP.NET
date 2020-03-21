<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/config.asp"-->
<!--#include file="../System/function.asp"-->
<%
  '如果学生还没有登陆
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'判断管理员是否开放了选课
	if canSelect = 0 then
	  Response.Write "<script>alert('对不起,现在没有开放选课功能!');location.href='../System/systemInfo.asp';</script>"
	else
	  if Now < selectStartTime or Now > selectEndTime then
	    Response.Write "<script>alert('请在选课时间内进行选课!');location.href='../System/systemInfo.asp';</script>"
	  end if
	end if
	'如果要求更新选课表
	if  Request("submit") <> "" then
	  '首先删除该学生该专业该学期的所有选修课程的选修课程
	  sqlString = "delete from [studentSelectCourseInfo] where studentNumber='" & session("studentNumber") & "' and courseNumber in (select courseNumber from [publicCourseInfo] where specialFieldNumber='" & GetSpecialFieldNumberByStudentNumber(session("studentNumber")) & "' and termId=" & termId & ")"
	  conn.Execute(sqlString)
	  '取得选修课的编号
	  courseNumbers = Request("courseNumbers")
	  if courseNumbers <> "" then
	    courseNumbers = Split(courseNumbers,",")
		  for each courseNumber in courseNumbers
		    sqlString = "insert into [studentSelectCourseInfo] values ('" & session("studentNumber") & "','" & Trim(courseNumber) & "')"
			  conn.Execute(sqlString)
		  next
	  end if
	  Response.Write "<script>alert('选课成功!');</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>学生选课</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="courseSelect.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=5 align="center">
		      <img src="../images/list.gif" width=14px height=14px>选课信息管理--&gt;学生选课
			 </td>
	   </tr>
		 <tr><td colspan=5><%=GetTermnameById(termId)%>&nbsp;&nbsp;选修课程信息</td></tr>
		 <%
		   sqlString = "select specialFieldNumber from [specialFieldInfo],[classInfo],[studentInfo] where [specialFieldInfo].specialFieldNumber = [classInfo].classSpecialFieldNumber and [classInfo].classNumber = [studentInfo].studentClassNumber and [studentInfo].studentNumber='" & Session("studentNumber") & "'"
			 set studentSpecialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
			 studentSpecialFieldInfoRs.Open sqlString,conn,1,1
			 specialFieldNumber = studentSpecialFieldInfoRs("specialFieldNumber")
			 studentSpecialFieldInfoRs.Close
			 '得到该专业该学期的所有选修课程信息
			 sqlString = "select * from [publicCourseInfo] where specialFieldNumber='" & specialFieldNumber & "' and termId=" & termId
			 set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
			 publicCourseInfoRs.Open sqlString,conn,1,1
			 if not publicCourseInfoRs.EOF then
			   Response.Write "<tr><td>选择</td><td>课程编号</td><td>课程名称</td><td>课程学分</td><td>上课老师</td></tr>"
			 else
			   Response.Write "<tr><td colspan=5>对不起,还没有选修课程信息</td></tr>"
			 end if
			 while not publicCourseInfoRs.EOF
			   Response.Write "<tr><td><input type=checkbox name=courseNumbers value='" & publicCourseInfoRs("courseNumber") & "'></td><td>" & publicCourseInfoRs("courseNumber") & "</td><td>" & publicCourseInfoRs("courseName") & "</td><td>" & publicCourseInfoRs("courseScore") & "</td><td>" & GetTeacherNameByPublicCourseNumber(publicCourseInfoRs("courseNumber")) & "</td></tr>"
			   publicCourseInfoRs.MoveNext
			wend
			publicCourseInfoRs.Close
		 %>
	  <tr bgcolor="#ffffff">
        <td height="30" colspan="5" align="center">
		      <input name="submit"  type="submit" value=" 选课 ">
				  <input type="reset" value=" 重新填写 ">
		    </td>
	 </tr>
	</form>
 </table>
 <table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=5 align="center">
		      <img src="../images/list.gif" width=14px height=14px><%=GetTermNameById(termId)%>&nbsp;你的选课信息
			 </td>
	   </tr>
		<%
		  sqlString = "select [studentSelectCourseInfo].courseNumber,[publicCourseInfo].courseName from [studentSelectCourseInfo],[publicCourseInfo] where [studentSelectCourseInfo].studentNumber='" & session("studentNumber") & "' and [studentSelectCourseInfo].courseNumber=[publicCourseInfo].courseNumber and [publicCourseInfo].termId=" & termId
		  set selectCourseRs = Server.CreateObject("ADODB.RecordSet")
		  selectCourseRs.Open sqlString,conn,1,1
		  if selectCourseRs.EOF then
		    Response.Write "<tr><td colspan=2 align=center>你还没有选课</td></tr>"
			end if
		  while not selectCourseRs.EOF
		    Response.Write "<tr><td>" & selectCourseRs("courseNumber") & "</td><td>" & selectCourseRs("courseName") & "</td></tr>"
	      selectCourseRs.MoveNext	 
		  wend
		%>
</table>
</body>
</html>