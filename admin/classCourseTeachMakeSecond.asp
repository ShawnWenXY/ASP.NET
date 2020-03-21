<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,courseNumber,courseName,classNumber,termId,className,termName
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得要排课的课程编号
	courseNumber = Request("courseNumber")
	sqlString = "select termId,classNumber,courseName from [classCourseInfo] where courseNumber='" & courseNumber & "'"
	set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseInfoRs.Open sqlString,conn,1,1
	'取得该课程的课程名称,班级编号和所在学期编号
	if not classCourseInfoRs.EOF then
	  courseName = classCourseInfoRs("courseName")
	  classNumber = classCourseInfoRs("classNumber")
	  termId = CInt(classCourseInfoRs("termId"))
	end if
	'根据班级编号取得班级名称
	className = GetClassNameByNumber(classNumber)
	'根据学期编号取得学期名称
	termName = GetTermnameById(termId)
%>
<HTML>
<HEAD>
	<Title>班级课程排课</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>排课信息管理--&gt;课程排课信息列表
			 </td>
	   </tr>
		<tr>
		  <td colspan=2 style="color:red;">
		  <%
		     Response.Write className & " " & termName & " " &  courseName & "课程的排课信息"
		  %>
		  </td>
		</tr>
		<%
		  '根据课程编号得到该课程的排课信息
		  sqlString = "select * from [classCourseTeach] where courseNumber='" & courseNumber & "'"
		  set classCourseTeachRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseTeachRs.Open sqlString,conn,1,1
		  if not classCourseTeachRs.EOF then
		    Response.Write "<tr><td>讲课教师</td><td>上课地点</td><td>上课时间</td><td>操作</td></tr>"
		  else
		    Response.Write "<tr><td align=center colspan=4>对不起,还没有该课程的排课信息!</td></tr>"
		  end if
		  while not classCourseTeachRs.EOF
		    Response.Write "<tr><td>" & GetTeacherNameByNumber(classCourseTeachRs("teacherNumber")) & "</td><td>" & classCourseTeachRs("teachClassRoom") & "</td><td>星期" & classCourseTeachRs("teachDay") & vbCrlf
			  Response.Write "</td><td><a href='classCourseTeachDetail.asp?teachId=" & classCourseTeachRs("teachId") & "'><img src='../images/edit.gif' height=12 width=12 border=0>详细</a>&nbsp;<a href='classCourseTeachDel.asp?teachId=" & classCourseTeachRs("teachId") & "' onclick=" & """" & "javascript:return confirm('真的确定删除此记录吗?')" & """" & "><img src='../images/delete.gif' height=12 width=12 border=0>删除</a></td></tr>" & vbCrlf
			  classCourseTeachRs.MoveNext
		  wend
		  classCourseTeachRs.Close
		%>
	  <tr><td colspan=4><input type="button" value="添加新的排课信息" onClick="javascript:location.href='classCourseTeachAdd.asp?courseNumber=<%=courseNumber%>'">
	      &nbsp;<input type="button" value="重新选择班级课程信息"  onClick="javascript:location.href='classCourseTeachMakeFirst.asp?termId=<%=termId%>&classNumber=<%=classNumber%>&submit=true'">
		  </td></tr>
	</table>
</BODY>
</HTML>


