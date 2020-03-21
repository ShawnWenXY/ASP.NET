<!--#include virtual="/Database/conn.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '取得删除的班级课程编号
  courseNumber = Request.QueryString("courseNumber")
  '首先删除该课程的成绩信息
  sqlString = "delete from [scoreInfo] where courseNumber='" & courseNumber & "' and isSelect=0"
  conn.Execute(sqlString)
  '然后删除该课程的上课信息
  sqlString = "delete from [classCourseTeach] where courseNumber='" & courseNumber & "'"
  conn.Execute(sqlString)
  '最后删除该班级课程的信息
  sqlString = "delete from [classCourseInfo] where courseNumber='" & courseNumber & "'"
  conn.Execute(sqlString)
  
  Response.Write "<script>alert('班级课程信息删除成功!');location.href='classCourseManage.asp';</script>"
%>