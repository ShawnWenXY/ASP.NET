<!--#include virtual="/DataBase/conn.asp"-->
<%

  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	dim studentNumber,sqlString
	'取得要删除学生的学号信息
	studentNumber = Request.QueryString("studentNumber")
	'删除该学生的选课信息
	sqlString = "delete from [studentSelectCourseInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	'删除该学生的成绩信息
	sqlString = "delete from [scoreInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	'最后删除该学生的记录信息
	sqlString = "delete from [studentInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	
	Response.Write "<script>alert('学生信息删除成功!');location.href='studentInfoManage.asp'</script>"
	
%>

