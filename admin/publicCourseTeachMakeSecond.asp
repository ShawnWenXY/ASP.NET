<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,courseNumber,courseName,specialFieldNumber,termId,specialFieldName,termName
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得要排课的课程编号
	courseNumber = Request("courseNumber")
	sqlString = "select termId,specialFieldNumber,courseName from [publicCourseInfo] where courseNumber='" & courseNumber & "'"
	set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseInfoRs.Open sqlString,conn,1,1
	'取得该课程的课程名称,专业编号和所在学期编号
	if not publicCourseInfoRs.EOF then
	  courseName = publicCourseInfoRs("courseName")
	  specialFieldNumber = publicCourseInfoRs("specialFieldNumber")
	  termId = CInt(publicCourseInfoRs("termId"))
	end if
	'根据专业编号取得专业名称
	specialFieldName = GetSpecialFieldNameByNumber(specialFieldNumber)
	'根据学期编号取得学期名称
	termName = GetTermnameById(termId)
%>
<HTML>
<HEAD>
	<Title>专业公选课程排课</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>排课信息管理--&gt;选修课程排课信息列表
			 </td>
	   </tr>
		<tr>
		  <td colspan=2 style="color:red;">
		  <%
		     Response.Write specialFieldName & " 专业 " & termName & " " &  courseName & "课程的排课信息"
		  %>
		  </td>
		</tr>
		<%
		  '根据课程编号得到该课程的排课信息
		  sqlString = "select * from [publicCourseTeach] where courseNumber='" & courseNumber & "'"
		  set publicCourseTeachRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseTeachRs.Open sqlString,conn,1,1
		  if not publicCourseTeachRs.EOF then
		    Response.Write "<tr><td>讲课教师</td><td>上课地点</td><td>上课时间</td><td>操作</td></tr>"
		  else
		    Response.Write "<tr><td align=center colspan=4>对不起,还没有该课程的排课信息!</td></tr>"
		  end if
		  while not publicCourseTeachRs.EOF
		    Response.Write "<tr><td>" & GetTeacherNameByNumber(publicCourseTeachRs("teacherNumber")) & "</td><td>" & publicCourseTeachRs("teachClassRoom") & "</td><td>星期" & publicCourseTeachRs("teachDay") & vbCrlf
			  Response.Write "</td><td><a href='publicCourseTeachDetail.asp?teachId=" & publicCourseTeachRs("teachId") & "'><img src='../images/edit.gif' height=12 width=12 border=0>详细</a>&nbsp;<a href='publicCourseTeachDel.asp?teachId=" & publicCourseTeachRs("teachId") & "' onclick=" & """" & "javascript:return confirm('真的确定删除此记录吗?')" & """" & "><img src='../images/delete.gif' height=12 width=12 border=0>删除</a></td></tr>" & vbCrlf
			  publicCourseTeachRs.MoveNext
		  wend
		  publicCourseTeachRs.Close
		%>
	  <tr><td colspan=4><input type="button" value="添加新的排课信息" onClick="javascript:location.href='publicCourseTeachAdd.asp?courseNumber=<%=courseNumber%>'">
	      &nbsp;<input type="button" value="重新选择班级课程信息"  onClick="javascript:location.href='publicCourseTeachMakeFirst.asp?termId=<%=termId%>&specialFieldNumber=<%=specialFieldNumber%>&submit=true'">
		  </td></tr>
	</table>
</BODY>
</HTML>


