<!--#include virtual="/DataBase/conn.asp"-->
<%
  dim teachId,sqlString,publicCourseTeachInfoRs,courseNumber
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得要删除上课信息的编号
	teachId = CInt(Request("teachId"))
	'取得关于上课信息的课程编号
	sqlString = "select courseNumber from [publicCourseTeach] where teachId=" & teachId
	set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseTeachInfoRs.Open sqlString,conn,1,1
	courseNumber = publicCourseTeachInfoRs("courseNumber")
	sqlString = "delete from [publicCourseTeach] where teachId=" & teachId
	conn.Execute(sqlString)
	Response.Write "<script>alert('上课信息删除成功!');location.href='publicCourseTeachMakeSecond.asp?courseNumber=" & courseNumber & "';</script>"
%>