<!--#include virtual="/DataBase/conn.asp"-->
<%
  dim teachId,sqlString,classCourseTeachInfoRs,courseNumber
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得要删除上课信息的编号
	teachId = CInt(Request("teachId"))
	'取得关于上课信息的课程编号
	sqlString = "select courseNumber from [classCourseTeach] where teachId=" & teachId
	set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseTeachInfoRs.Open sqlString,conn,1,1
	courseNumber = classCourseTeachInfoRs("courseNumber")
	sqlString = "delete from [classCourseTeach] where teachId=" & teachId
	conn.Execute(sqlString)
	Response.Write "<script>alert('上课信息删除成功!');location.href='classCourseTeachMakeSecond.asp?courseNumber=" & courseNumber & "';</script>"
%>