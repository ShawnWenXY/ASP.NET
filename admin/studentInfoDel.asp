<!--#include virtual="/DataBase/conn.asp"-->
<%

  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	dim studentNumber,sqlString
	'ȡ��Ҫɾ��ѧ����ѧ����Ϣ
	studentNumber = Request.QueryString("studentNumber")
	'ɾ����ѧ����ѡ����Ϣ
	sqlString = "delete from [studentSelectCourseInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	'ɾ����ѧ���ĳɼ���Ϣ
	sqlString = "delete from [scoreInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	'���ɾ����ѧ���ļ�¼��Ϣ
	sqlString = "delete from [studentInfo] where studentNumber='" & studentNumber & "'"
	conn.Execute(sqlString)
	
	Response.Write "<script>alert('ѧ����Ϣɾ���ɹ�!');location.href='studentInfoManage.asp'</script>"
	
%>

