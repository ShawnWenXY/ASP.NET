<!--#include virtual="/DataBase/conn.asp"-->
<%
  dim teachId,sqlString,classCourseTeachInfoRs,courseNumber
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ��Ҫɾ���Ͽ���Ϣ�ı��
	teachId = CInt(Request("teachId"))
	'ȡ�ù����Ͽ���Ϣ�Ŀγ̱��
	sqlString = "select courseNumber from [classCourseTeach] where teachId=" & teachId
	set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseTeachInfoRs.Open sqlString,conn,1,1
	courseNumber = classCourseTeachInfoRs("courseNumber")
	sqlString = "delete from [classCourseTeach] where teachId=" & teachId
	conn.Execute(sqlString)
	Response.Write "<script>alert('�Ͽ���Ϣɾ���ɹ�!');location.href='classCourseTeachMakeSecond.asp?courseNumber=" & courseNumber & "';</script>"
%>