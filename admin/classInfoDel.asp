<!--#include virtual="/Database/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  'ȡ��ɾ���İ༶���
  classNumber = Request.QueryString("classNumber")
  '��ѯ�༶���Ƿ񻹴���ѧ����Ϣ�����������ʾɾ�����ɹ�
  sqlString = "select * from [studentInfo] where studentClassNumber='" & classNumber & "'"
  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
  studentInfoRs.Open sqlString,conn,1,1
  if not studentInfoRs.EOF then
    Response.Write "<script>alert('�ð༶�»�����ѧ����Ϣ�����Ƚ�ѧ����Ϣɾ��');location.href='classInfoManage.asp';</script>"
	else
	  '����ɾ���༶��Ϣ
	  sqlString = "delete from [classInfo] where classNumber='" & classNumber & "'"
	  conn.Execute(sqlString)
	  'Ȼ��ɾ���ð༶�Ŀγ���Ϣ
	  sqlString = "delete from [classCourseInfo] where classNumber='" & classNumber & "'"
	  conn.Execute(sqlString)
	  'ע��:�˴�û��ɾ���ɼ���ѡ�εĴ�������Ϊ��ѧ����Ϣɾ��ʱ�Ѿ�����ɾ����
	  Response.Write("<script>alert('�༶��Ϣɾ���ɹ�!');location.href='classInfoManage.asp';</script>")
	end if
%>