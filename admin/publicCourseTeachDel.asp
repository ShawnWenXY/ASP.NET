<!--#include virtual="/DataBase/conn.asp"-->
<%
  dim teachId,sqlString,publicCourseTeachInfoRs,courseNumber
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ��Ҫɾ���Ͽ���Ϣ�ı��
	teachId = CInt(Request("teachId"))
	'ȡ�ù����Ͽ���Ϣ�Ŀγ̱��
	sqlString = "select courseNumber from [publicCourseTeach] where teachId=" & teachId
	set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseTeachInfoRs.Open sqlString,conn,1,1
	courseNumber = publicCourseTeachInfoRs("courseNumber")
	sqlString = "delete from [publicCourseTeach] where teachId=" & teachId
	conn.Execute(sqlString)
	Response.Write "<script>alert('�Ͽ���Ϣɾ���ɹ�!');location.href='publicCourseTeachMakeSecond.asp?courseNumber=" & courseNumber & "';</script>"
%>