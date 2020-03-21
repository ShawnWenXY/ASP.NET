<!--#include file="Database/conn.asp"-->
<!--#include file="System/md5.asp"-->
<%
	'取得用户名,密码,身份信息
  dim username,password,identity,sqlString
  username = Trim(CStr(Request("username")))
  password = Trim(CStr(Request("password")))
  identity = Trim(CStr(Request("identity")))
  '创建与登陆信息相关的结果记录集
  set rsLogin = Server.CreateObject("adodb.recordset")
  '如果是学生身份
  if identity = "student" then
    sqlString = "select * from [studentInfo] where studentNumber='" & username & "' and studentPassword='" & md5(password) & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '如果学生学号和密码都输入正确
	  if not rsLogin.EOF then
	    Session("studentNumber") = username
		  Response.Redirect "Student/index.asp"
		  Response.End
		'如果该学号和密码的记录不存在
	  else
	    Response.Write "<script>alert('学号或密码输入错误!');location.href='login.asp';</script>"
		  Response.End
		end if
	'如果是教师身份
	elseif identity = "teacher" then
	  sqlString = "select * from [teacherInfo] where teacherNumber='" & username & "' and teacherPassword='" & md5(password) & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '如果教师教工号和密码都输入正确
	  if not rsLogin.EOF then
	    Session("teacherNumber") = username
		  Response.Redirect "Teacher/index.asp"
		'如果不存在该教工号的教师信息
		else
		  Response.Write "<script>alert('教师登陆帐号或密码错误!');location.href='login.asp';</script>"
		  Response.End
		end if
	'如果是管理员身份登陆
	else
	  sqlString = "select * from [admin] where adminUsername='" & username & "' and adminPassword='" & password & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '如果管理员帐号和密码都输入正确
	  if not rsLogin.EOF then
	    Session("adminUsername") = username
		  Response.Redirect "Admin/index.asp"
		'如果管理员的信息不存在
		else
		  Response.Write "<script>alert('该管理员的信息不存在!');location.href='login.asp';</script>"
		end if
	end if
%>
