<!--#include virtual="/Database/conn.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得要删除的专业编号
	specialFieldNumber = Trim(Request("specialFieldNumber"))
	'检查该专业下是否存在班级信息,如果还存在班级需要先删除班级后才能执行专业信息的删除操作
	sqlString = "select * from [classInfo] where classSpecialFieldNumber='" & specialFieldNumber & "'"
	set classInfoRs = Server.CreateObject("ADODB.RecordSet")
	classInfoRs.Open sqlString,conn,1,1
	if not classInfoRs.EOF then
	  Response "<script>alert('该专业下还存在班级,请先删除班级信息!');location.href='specialFieldInfoManage.asp';</script>"	
	else  '如果不存在班级信息了
	  '先执行该专业选修课程的删除操作(每个班级的必修课程在班级信息删除操作时执行删除操作)
	  sqlString = "delete from [publicCourseInfo] where specialFieldNumber='" & specialFieldNumber & "'"
	  conn.Execute(sqlString)
	  '再执行该专业信息的删除
	  sqlString = "delete from [specialFieldInfo] where specialFieldNumber='" & specialFieldNumber & "'"
	  conn.Execute(sqlString)
	  Response.Write "<script>alert('专业信息删除成功!');location.href='specialFieldInfoManage.asp';</script>"
	end if
%>