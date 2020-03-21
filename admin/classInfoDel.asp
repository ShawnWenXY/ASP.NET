<!--#include virtual="/Database/conn.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '取得删除的班级编号
  classNumber = Request.QueryString("classNumber")
  '查询班级下是否还存在学生信息，如果存在提示删除不成功
  sqlString = "select * from [studentInfo] where studentClassNumber='" & classNumber & "'"
  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
  studentInfoRs.Open sqlString,conn,1,1
  if not studentInfoRs.EOF then
    Response.Write "<script>alert('该班级下还存在学生信息，请先将学生信息删除');location.href='classInfoManage.asp';</script>"
	else
	  '首先删除班级信息
	  sqlString = "delete from [classInfo] where classNumber='" & classNumber & "'"
	  conn.Execute(sqlString)
	  '然后删除该班级的课程信息
	  sqlString = "delete from [classCourseInfo] where classNumber='" & classNumber & "'"
	  conn.Execute(sqlString)
	  '注意:此处没有删除成绩和选课的代码是因为在学生信息删除时已经将其删除了
	  Response.Write("<script>alert('班级信息删除成功!');location.href='classInfoManage.asp';</script>")
	end if
%>